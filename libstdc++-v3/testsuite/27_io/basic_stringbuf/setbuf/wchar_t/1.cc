// 981208 bkoz test functionality of basic_stringbuf for char_type == wchar_t

// Copyright (C) 1997, 1998, 1999, 2000, 2001, 2002, 2003, 2004
// Free Software Foundation, Inc.
//
// This file is part of the GNU ISO C++ Library.  This library is free
// software; you can redistribute it and/or modify it under the
// terms of the GNU General Public License as published by the
// Free Software Foundation; either version 2, or (at your option)
// any later version.

// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License along
// with this library; see the file COPYING.  If not, write to the Free
// Software Foundation, 59 Temple Place - Suite 330, Boston, MA 02111-1307,
// USA.

#include <sstream>
#include <testsuite_hooks.h>

std::wstring str_01(L"mykonos. . . or what?");
std::wstring str_02(L"paris, or sainte-maxime?");
std::wstring str_03;
std::wstringbuf strb_01(str_01);
std::wstringbuf strb_02(str_02, std::ios_base::in);
std::wstringbuf strb_03(str_03, std::ios_base::out);

// test overloaded virtual functions
void test04() 
{
  bool test __attribute__((unused)) = true;
  std::wstring 		str_tmp;
  std::wstringbuf 		strb_tmp;
  typedef std::wstringbuf::int_type int_type;
  typedef std::wstringbuf::traits_type traits_type;
  typedef std::wstringbuf::pos_type pos_type;
  typedef std::wstringbuf::off_type off_type;

  // PUT
  strb_03.str(str_01); //reset
  
  // BUFFER MANAGEMENT & POSITIONING
  // setbuf
  // pubsetbuf(char_type* s, streamsize n)
  str_tmp = std::wstring(L"naaaah, go to cebu");
  strb_01.pubsetbuf(const_cast<wchar_t*> (str_tmp.c_str()), str_tmp.size());
  VERIFY( strb_01.str() == str_tmp );
  strb_01.pubsetbuf(0,0);
  VERIFY( strb_01.str() == str_tmp );
}

int main()
{
  test04();
  return 0;
}



// more candy!!!
