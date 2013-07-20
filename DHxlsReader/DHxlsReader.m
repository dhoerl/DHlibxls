/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
 *
 * This file is part of DHlibxls -- permitting code to read Excel(TM) files.
 *
 * Copyright 2012 David Hoerl All Rights Reserved.
 *
 * Redistribution and use in source and binary forms, with or without modification, are
 * permitted provided that the following conditions are met:
 * 
 *    1. Redistributions of source code must retain the above copyright notice, this list of
 *       conditions and the following disclaimer.
 * 
 *    2. Redistributions in binary form must reproduce the above copyright notice, this list
 *       of conditions and the following disclaimer in the documentation and/or other materials
 *       provided with the distribution.
 *
 *    3. Only the ObjectiveC code is under this license - separate license may apply to libxls source.
 *       Read the documents and source headers in the libxls files which you downloaded earlier.
 * 
 * THIS SOFTWARE IS PROVIDED BY David Hoerl ''AS IS'' AND ANY EXPRESS OR IMPLIED
 * WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND
 * FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL David Hoerl OR
 * CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 * SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
 * ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
 * NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF
 * ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */


#import "DHxlsReader.h"

#import "DHcell-Private.h"
#import "xls.h"

#if ! __has_feature(objc_arc)
#error THIS CODE MUST BE COMPILED WITH ARC ENABLED!
#endif
 

@interface DHxlsReader ()

- (void)setWorkBook:(xlsWorkBook *)wb;

- (void)openSheet:(uint32_t)sheetNum;
- (void)formatContent:(DHcell *)content withCell:(xlsCell *)cell;

@end

@implementation DHxlsReader
{
	xlsWorkBook			*_workBook;
	uint32_t			_numSheets;
	uint32_t			_activeWorkSheetID;		// keep last one active
	xlsWorkSheet		*_activeWorkSheet;		// keep last one active
	xlsSummaryInfo		*_summary;
	
	BOOL				_iterating;
	uint32_t			_lastRowIndex;
	uint32_t			_lastColIndex;
	
	NSStringEncoding	_encoding;
}

+ (DHxlsReader *)xlsReaderWithPath:(NSString *)filePath
{
	DHxlsReader			*reader;
	xlsWorkBook			*workBook;

	// NSLog(@"sizeof FORMULA=%zd LABELSST=%zd", sizeof(FORMULA), sizeof(LABELSST) );
	const char *file = [filePath cStringUsingEncoding:NSUTF8StringEncoding];
	if((workBook = xls_open(file, "UTF-8"))) {
		reader = [DHxlsReader new];
		[reader setWorkBook:workBook];
	}
	return reader;
}

+ (DHxlsReader *)xlsReaderFromFile:(NSString *)filePath
{
	return [self xlsReaderWithPath:filePath];
}


- (id)init
{
	if((self = [super init])) {
		_activeWorkSheetID = DHWorkSheetNotFound;
		_encoding = NSUTF8StringEncoding;
	}
	return self;
}
- (void)dealloc
{
	xls_close_summaryInfo(_summary);
	xls_close_WS(_activeWorkSheet);
	xls_close_WB(_workBook);
}

- (void)setWorkBook:(xlsWorkBook *)wb
{
	_workBook = wb;
	xls_parseWorkBook(_workBook);
	_numSheets = _workBook->sheets.count;
	_summary = xls_summaryInfo(_workBook);
}

- (NSString *)libaryVersion
{
	return [NSString stringWithCString:xls_getVersion() encoding:NSASCIIStringEncoding];
}

// Sheet Information
- (uint32_t)numberOfSheets
{
	return _numSheets;
}

- (NSString *)sheetNameAtIndex:(uint32_t)idx
{
	return idx < _numSheets ? [NSString stringWithCString:(char *)_workBook->sheets.sheet[idx].name encoding:_encoding] : nil;
}

- (uint16_t)rowsForSheetAtIndex:(uint32_t)idx
{
    [self openSheet:idx];
    NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
	return idx < _numSheets ? numRows : 0;
}

- (BOOL)isSheetVisibleAtIndex:(NSUInteger)idx
{
	return idx < _numSheets ? (BOOL)_workBook->sheets.sheet[idx].visibility : NO;
}

- (void)openSheet:(uint32_t)sheetNum
{	
	if(sheetNum >= _numSheets) {
		_iterating = true;
		_lastColIndex = UINT32_MAX;
		_lastRowIndex = UINT32_MAX;
	} else
	if(sheetNum != _activeWorkSheetID) {
		_activeWorkSheetID = sheetNum;
		xls_close_WS(_activeWorkSheet);
		_activeWorkSheet = xls_getWorkSheet(_workBook, sheetNum);
		xls_parseWorkSheet(_activeWorkSheet);
	}
}

- (uint16_t)numberOfRowsInSheet:(uint32_t)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastrow + 1;
}

- (uint16_t)numberOfColsInSheet:(uint32_t)sheetIndex
{
    [self openSheet:sheetIndex];
    return _activeWorkSheet->rows.lastcol + 1;
}

// Random Access
- (DHcell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row col:(uint16_t)col
{
	DHcell *content = [DHcell blankCell];
	
	assert(row && col);

	[self startIterator:DHWorkSheetNotFound];
	[self openSheet:sheetNum];
	
	--row, --col;
	
	NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
	NSUInteger numCols = _activeWorkSheet->rows.lastcol + 1;

	for (NSUInteger t=0; t<numRows; t++)
	{
		xlsRow *rowP = &_activeWorkSheet->rows.row[t];
		for (NSUInteger tt=0; tt<numCols; tt++)
		{
			xlsCell	*cell = &rowP->cells.cell[tt];
			// NSLog(@"Looking for %d:%d:%d - testing %d:%d Type: 0x%4.4x  [t=%d tt=%d]", sheetNum, row, col, cell->row, cell->col, cell->id, t, tt);
			if(cell->row < row) break;
			if(cell->row > row) return content;
			
			if(cell->id == 0x201) continue;	// "Blank" filler cell created by libxls
			
			if(cell->col == col) {
				[self formatContent:content withCell:cell];
				return content;
			}
		}
	}
	
	return content;
}

- (DHcell *)cellInWorkSheetIndex:(uint32_t)sheetNum row:(uint16_t)row colStr:(char *)colStr
{
	if(strlen(colStr) > 2 || strlen(colStr) == 0) return [DHcell blankCell];

	NSInteger col = colStr[0] - 'A';
	if(col < 0 || col >= 26) return [DHcell blankCell];
	char c = colStr[1];
	if(c) {
		col *= 26;
		NSInteger col2 = c - 'A';
		if(col2 < 0 || col2 >= 26) return [DHcell blankCell];
		col += col2;
	}
	col += 1;

	return [self cellInWorkSheetIndex:sheetNum row:row col:(uint16_t)col];
}

// Iterate through all cells
- (void)startIterator:(uint32_t)sheetNum
{
	if(sheetNum != DHWorkSheetNotFound) {
		[self openSheet:sheetNum];
		_iterating = true;
		_lastColIndex = 0;
		_lastRowIndex = 0;
	} else {
		_iterating = false;
	}
}

- (DHcell *)nextCell
{
	DHcell *content = [DHcell blankCell];

	if(!_iterating) return nil;
	
	NSUInteger numRows = _activeWorkSheet->rows.lastrow + 1;
	NSUInteger numCols = _activeWorkSheet->rows.lastcol + 1;

	if(_lastRowIndex >= numRows) return content;
	
	for (NSUInteger t=_lastRowIndex; t<numRows; t++)
	{
		xlsRow *rowP = &_activeWorkSheet->rows.row[t];
		for (uint32_t tt=_lastColIndex; tt<numCols; tt++)
		{
			xlsCell	*cell = &rowP->cells.cell[tt];
			
			if(cell->id == 0x201) continue;
			_lastColIndex = tt + 1;
			[self formatContent:content withCell:cell];
			return content;
		}
		++_lastRowIndex;
		_lastColIndex = 0;
	}
	// don't make iterator false - user can keep asking for cells, they all just be blank ones though
	return content;
}

- (void)formatContent:(DHcell *)content withCell:(xlsCell *)cell
{
	NSUInteger col = cell->col;

	content.row = cell->row + 1;
	
	{
		content.col = col + 1;
		char colStr[3];
		if(col < 26) {
			colStr[0] = 'A' + (char)col;
			colStr[1] = '\0';
		} else {
			colStr[0] = 'A' + (char)(col/26);
			colStr[1] = 'A' + (char)(col%26);
		}
		colStr[2] = '\0';
		[content setColStr:colStr];
	}

	switch(cell->id) {
    case 0x0006:	//FORMULA
		// test for formula, if
        if(cell->l == 0) {
			content.type = cellFloat;
			content.val = [NSNumber numberWithDouble:cell->d];
		} else {
			if(!strcmp((char *)cell->str, "bool")) {
				BOOL b = (BOOL)cell->d;
				content.type = cellBool;
				content.val = [NSNumber numberWithBool:b];
				content.str = b ? @"YES" : @"NO";
			} else
			if(!strcmp((char *)cell->str, "error")) {
				// FIXME: Why do we convert the double cell->d to NSInteger?
				NSInteger err = (NSInteger)cell->d;
				content.type = cellError;
				content.val = [NSNumber numberWithInteger:err];
				content.str = [NSString stringWithFormat:@"%ld", (long)err];
			} else {
				content.type = cellString;
			}
		}
        break;
    case 0x00FD:	//LABELSST
    case 0x0204:	//LABEL
		content.type = cellString;
		content.val = [NSNumber numberWithLong:cell->l];	// possible numeric conversion done for you
		break;
    case 0x0203:	//NUMBER
    case 0x027E:	//RK
		content.type = cellFloat;
		content.val = [NSNumber numberWithDouble:cell->d];
        break;
    default:
		content.type = cellUnknown;
        break;
    }
	
	if(!content.str) {
		content.str = [NSString stringWithCString:(char *)cell->str encoding:NSUTF8StringEncoding];
	}
	// NSLog(@"GOING TO PRINT STRING");
	// NSLog(@"Cell creator: t=%d num=%@ str=%@", content.type, content.val, content.str);
}

// Summary Information
- (NSString *)appName		{ return _summary->appName	? [NSString stringWithCString:(char *)_summary->appName		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)author		{ return _summary->author	? [NSString stringWithCString:(char *)_summary->author		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)category		{ return _summary->category	? [NSString stringWithCString:(char *)_summary->category		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)comment		{ return _summary->comment	? [NSString stringWithCString:(char *)_summary->comment		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)company		{ return _summary->company	? [NSString stringWithCString:(char *)_summary->company		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)keywords		{ return _summary->keywords	? [NSString stringWithCString:(char *)_summary->keywords		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)lastAuthor	{ return _summary->lastAuthor? [NSString stringWithCString:(char *)_summary->lastAuthor	encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)manager		{ return _summary->manager	? [NSString stringWithCString:(char *)_summary->manager		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)subject		{ return _summary->subject	? [NSString stringWithCString:(char *)_summary->subject		encoding:NSUTF8StringEncoding] : @""; }
- (NSString *)title			{ return _summary->title		? [NSString stringWithCString:(char *)_summary->title		encoding:NSUTF8StringEncoding] : @""; }

@end
