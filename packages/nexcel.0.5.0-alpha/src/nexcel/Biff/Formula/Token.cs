/// <summary>******************************************************************
/// 
/// Copyright (C) 2005  Stefano Franco
///
/// Based on JExcelAPI by Andrew Khan.
/// 
/// This library is free software; you can redistribute it and/or
/// modify it under the terms of the GNU Lesser General Public
/// License as published by the Free Software Foundation; either
/// version 2.1 of the License, or (at your option) any later version.
/// 
/// This library is distributed in the hope that it will be useful,
/// but WITHOUT ANY WARRANTY; without even the implied warranty of
/// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
/// Lesser General Public License for more details.
/// 
/// You should have received a copy of the GNU Lesser General Public
/// License along with this library; if not, write to the Free Software
/// Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
/// *************************************************************************
/// </summary>
using System;
using System.Collections;
namespace NExcel.Biff.Formula
{
	
	/// <summary> An enumeration detailing the Excel parsed tokens
	/// A particular token may be associated with more than one token code
	/// </summary>
	class Token
	{
		/// <summary> Gets the token code for the specified token
		/// 
		/// </summary>
		/// <returns> the token code.  This is the first item in the array
		/// </returns>
		virtual public sbyte Code
		{
			get
			{
				return (sbyte) Value[0];
			}
			
		}
		/// <summary> Gets an alternative token code for the specified token
		/// Used for certain types of volatile function
		/// 
		/// </summary>
		/// <returns> the token code
		/// </returns>
		virtual public sbyte Code2
		{
			get
			{
				return (sbyte) (Value.Length > 0?Value[1]:Value[0]);
			}
			
		}
		/// <summary> The array of values which apply to this token</summary>
				public int[] Value;
		
		/// <summary> All available tokens, keyed on value</summary>
				private static Hashtable tokens;
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Token(int v)
		{
			Value = new int[]{v};
			
			tokens[v] =  this;
		}
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Token(int v1, int v2)
		{
			Value = new int[]{v1, v2};
			
			tokens[v1] =  this;
			tokens[v2] =  this;
		}
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Token(int v1, int v2, int v3)
		{
			Value = new int[]{v1, v2, v3};
			
			tokens[v1] =  this;
			tokens[v2] =  this;
			tokens[v3] =  this;
		}
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Token(int v1, int v2, int v3, int v4)
		{
			Value = new int[]{v1, v2, v3, v4};
			
			tokens[v1] =  this;
			tokens[v2] =  this;
			tokens[v3] =  this;
			tokens[v4] =  this;
		}
		
		/// <summary> Constructor
		/// Sets the token value and adds this token to the array of all token
		/// 
		/// </summary>
		/// <param name="v">the biff code for the token
		/// </param>
		private Token(int v1, int v2, int v3, int v4, int v5)
		{
			Value = new int[]{v1, v2, v3, v4, v5};
			
			tokens[v1] =  this;
			tokens[v2] =  this;
			tokens[v3] =  this;
			tokens[v4] =  this;
			tokens[v5] =  this;
		}
		
		/// <summary> Gets the type object from its integer value</summary>
		public static Token getToken(int v)
		{
			Token t = (Token) tokens[v];
			
			return t != null?t:UNKNOWN;
		}
		
		// Operands
		public static readonly Token REF;
		public static readonly Token REF3D;
		public static readonly Token MISSING_ARG;
		public static readonly Token STRING;
		public static readonly Token BOOL;
		public static readonly Token INTEGER;
		public static readonly Token DOUBLE;
		public static readonly Token REFV;
		public static readonly Token AREAV;
		public static readonly Token AREA;
		public static readonly Token NAMED_RANGE;
		public static readonly Token NAME;
		public static readonly Token AREA3D;
		
		// Unary Operators
		public static readonly Token UNARY_PLUS;
		public static readonly Token UNARY_MINUS;
		public static readonly Token PERCENT;
		public static readonly Token PARENTHESIS;
		
		// Binary Operators
		
		public static readonly Token ADD;
		
		public static readonly Token SUBTRACT;
		public static readonly Token MULTIPLY;
		public static readonly Token DIVIDE;
		public static readonly Token POWER;
		public static readonly Token CONCAT;
		public static readonly Token LESS_THAN;
		public static readonly Token LESS_EQUAL;
		public static readonly Token EQUAL;
		public static readonly Token GREATER_EQUAL;
		public static readonly Token GREATER_THAN;
		public static readonly Token NOT_EQUAL;
		public static readonly Token RANGE;
		
		// Functions
		public static readonly Token FUNCTION;
		public static readonly Token FUNCTIONVARARG;
		
		// Control
		public static readonly Token ATTRIBUTE;
		public static readonly Token MEM_FUNC;
		
		// Unknown token
		public static readonly Token UNKNOWN;


		static Token()
		{
			tokens = new Hashtable(20);

		// Operands
			REF = new Token(0x44, 0x24, 0x64);
			REF3D = new Token(0x5a, 0x3a, 0x7a);
			MISSING_ARG = new Token(0x16);
			STRING = new Token(0x17);
			BOOL = new Token(0x1d);
			INTEGER = new Token(0x1e);
			DOUBLE = new Token(0x1f);
			REFV = new Token(0x2c, 0x4c);
			AREAV = new Token(0x2d, 0x4d);
			AREA = new Token(0x25, 0x65, 0x45);
			NAMED_RANGE = new Token(0x43, 0x23);
			NAME = new Token(0x39);
			AREA3D = new Token(0x3b);
		
			// Unary Operators
			UNARY_PLUS = new Token(0x12);
			UNARY_MINUS = new Token(0x13);
			PERCENT = new Token(0x14);
			PARENTHESIS = new Token(0x15);
		
			// Binary Operators
		
			ADD = new Token(0x3);
		
			SUBTRACT = new Token(0x4);
			MULTIPLY = new Token(0x5);
			DIVIDE = new Token(0x6);
			POWER = new Token(0x7);
			CONCAT = new Token(0x8);
			LESS_THAN = new Token(0x9);
			LESS_EQUAL = new Token(0xa);
			EQUAL = new Token(0xb);
			GREATER_EQUAL = new Token(0xc);
			GREATER_THAN = new Token(0xd);
			NOT_EQUAL = new Token(0xe);
			RANGE = new Token(0x11);
		
			// Functions
			FUNCTION = new Token(0x41, 0x21, 0x61);
			FUNCTIONVARARG = new Token(0x42, 0x22, 0x62);
		
			// Control
			ATTRIBUTE = new Token(0x19);
			MEM_FUNC = new Token(0x29);
		
			// Unknown token
			UNKNOWN = new Token(0xffff);

		}
	}
}
