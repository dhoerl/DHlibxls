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

#import "DHcell-Private.h"

#if ! __has_feature(objc_arc)
#error THIS CODE MUST BE COMPILED WITH ARC ENABLED!
#endif
 

@implementation DHcell
{
	char colString[3];
}
@synthesize row;
@synthesize type;
@dynamic colStr;
@synthesize col;
@synthesize str;
@synthesize val;

+ (DHcell *)blankCell
{
	return [DHcell new];
}

- (char *)colStr
{
	return colString;
}
- (void)setColStr:(char *)colS
{
	colString[0] = colS[0];
	colString[1] = colS[1];
	colString[2] = '\0';
}

- (void)show
{
	NSLog(@"%@", [self dump]);
}

- (NSString *)dump
{
	NSMutableString *s = [NSMutableString stringWithCapacity:128];
	
	const char *name;
	switch(type) {
	case cellBlank:		name = "cellBlank";		break;
	case cellString:	name = "cellString";	break;
	case cellInteger:	name = "cellInteger";	break;
	case cellFloat:		name = "cellFloat";		break;
	case cellBool:		name = "cellBool";		break;
	case cellError:		name = "cellError";		break;
	default:			name = "cellUnknown";	break;
	}

	[s appendString:@"====================\n"];
	[s appendFormat:@"CellType: %s row=%u col=%s/%u\n", name, row, colString, col];
	[s appendFormat:@"   string:    %@\n", str];
	
	switch(type) {
	case cellInteger:	[s appendFormat:@"     long:    %ld\n",	[val longValue]];	break;
	case cellFloat:		[s appendFormat:@"    float:    %lf\n",	[val doubleValue]];	break;
	case cellBool:		[s appendFormat:@"     bool:    %d\n",	[val boolValue]];	break;
	case cellError:		[s appendFormat:@"    error:    %ld\n",	[val longValue]];	break;
	default: break;
	}
	return s;
}

@end
