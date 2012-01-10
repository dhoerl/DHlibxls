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


#import "DHcell.h"

@interface DHxlsReader : NSObject

+ (DHxlsReader *)xlsReaderFromFile:(NSString *)filePath;

- (NSString *)libaryVersion;

// Sheet Information
- (NSUInteger)numberOfSheets;
- (NSString *)sheetNameAtIndex:(NSUInteger)index;
- (BOOL)isSheetVisibleAtIndex:(NSUInteger)index;

// Random Access
- (DHcell *)cellInWorkSheetIndex:(NSUInteger)sheetNum row:(uint16_t)row col:(uint16_t)col;		// uses 1 based indexing!
- (DHcell *)cellInWorkSheetIndex:(NSUInteger)sheetNum row:(uint16_t)row colStr:(char *)col;		// "A"...."Z" "AA"..."ZZ"

// Iterate through all cells
- (void)startIterator:(NSUInteger)sheetNum;
- (DHcell *)nextCell;

// Summary Information
- (NSString *)appName;
- (NSString *)author;
- (NSString *)category;
- (NSString *)comment;
- (NSString *)company;
- (NSString *)keywords;
- (NSString *)lastAuthor;
- (NSString *)manager;
- (NSString *)subject;
- (NSString *)title;

@end
