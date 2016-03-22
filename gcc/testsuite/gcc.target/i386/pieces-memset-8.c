/* { dg-do compile } */
/* { dg-options "-O2 -mno-avx2 -mavx -mtune=generic" } */

extern char *dst;

void
foo (int x)
{
  __builtin_memset (dst, x, 17);
}

/* { dg-final { scan-assembler-times "vmovups\[ \\t\]+\[^\n\]*%xmm" 1 } } */
/* No need to dynamically realign the stack here.  */
/* { dg-final { scan-assembler-not "and\[^\n\r]*%\[re\]sp" } } */
/* Nor use a frame pointer.  */
/* { dg-final { scan-assembler-not "%\[re\]bp" } } */
