/* GCC MELT GENERATED FILE warmelt-normal+04.c - DO NOT EDIT */
/* secondary MELT generated C file of rank #4 */
#include "melt-run.h"


/* used hash from melt-run.h when compiling this file: */
MELT_EXTERN const char meltrun_used_md5_melt_f4[] =
  MELT_RUN_HASHMD5 /* from melt-run.h */ ;


/**** warmelt-normal+04.c declarations ****/
/**** MELT GENERATED DECLARATIONS for warmelt-normal ** DO NOT EDIT ; see gcc-melt.org ****/

/****++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
***
    Copyright 2008 - 2013 Free Software Foundation, Inc.
    Contributed by Basile Starynkevitch <basile@starynkevitch.net>

    This file is part of GCC.

    GCC is free software; you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation; either version 3, or (at your option)
    any later version.

    GCC is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with GCC; see the file COPYING3.  If not see
    <http://www.gnu.org/licenses/>.
***
----------------------------------------------------------------****/


/** ordinary MELT module meltbuild-sources/warmelt-normal**/
#define MELT_HAS_INITIAL_ENVIRONMENT 1	/*usualmodule */


/**** no MELT module variables ****/

/*** 3 MELT called hook declarations ***/

/*declare MELT called hook #0 HOOK_FRESH_ENVIRONMENT_REFERENCE_MAKER **/
MELT_EXTERN melt_ptr_t
melthook_HOOK_FRESH_ENVIRONMENT_REFERENCE_MAKER (melt_ptr_t melthookdata,
						 melt_ptr_t meltinp0_PREVENV,
						 const char
						 *meltinp1_MODULNAME);

/*declare MELT called hook #1 HOOK_SYMBOL_IMPORTER **/
MELT_EXTERN melt_ptr_t melthook_HOOK_SYMBOL_IMPORTER (melt_ptr_t melthookdata,
						      const char
						      *meltinp0_SYMNAMESTR,
						      const char
						      *meltinp1_MODULENAMESTR,
						      melt_ptr_t
						      meltinp2_PARENV);

/*declare MELT called hook #2 HOOK_VALUE_EXPORTER **/
MELT_EXTERN void melthook_HOOK_VALUE_EXPORTER (melt_ptr_t melthookdata,
					       melt_ptr_t meltinp0_SYM,
					       melt_ptr_t meltinp1_VAL,
					       melt_ptr_t meltinp2_CONTENV);

/*** end of 3 MELT called hook declarations ***/

/*** no extra MELT c-headers ***/




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_1_WARMELTmiNORMAL_NORMAL_IMPORT (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_2_WARMELTmiNORMAL_ADD_NCTX_DATA (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_3_WARMELTmiNORMAL_FILL_INITIAL_PREDEFMAP (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_4_WARMELTmiNORMAL_REGISTER_LITERAL_VALUE (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_5_WARMELTmiNORMAL_CREATE_NORMCONTEXT (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_6_WARMELTmiNORMAL_CREATE_NORMAL_EXTENDING_CONTEXT (meltclosure_ptr_t
							    meltclosp_,
							    melt_ptr_t
							    meltfirstargp_,
							    const
							    melt_argdescr_cell_t
							    meltxargdescr_[],
							    union meltparam_un
							    *meltxargtab_,
							    const
							    melt_argdescr_cell_t
							    meltxresdescr_[],
							    union meltparam_un
							    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_7_WARMELTmiNORMAL_NORMEXP_IDENTICAL (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_8_WARMELTmiNORMAL_NORMEXP_NULL (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_9_WARMELTmiNORMAL_NORMEXP_ANY_OBJECT (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_10_WARMELTmiNORMAL_NORMEXP_ANY_VALUE (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_11_WARMELTmiNORMAL_NORMEXP_SRC_CATCHALL (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_12_WARMELTmiNORMAL_NORMEXP_LAZYMACROEXP (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_13_WARMELTmiNORMAL_GECTYP_ANYRECV (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_14_WARMELTmiNORMAL_GECTYP_ROOT (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_15_WARMELTmiNORMAL_GECTYP_INTEGER (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_16_WARMELTmiNORMAL_GECTYP_STRING (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_17_WARMELTmiNORMAL_NORMALIZE_TUPLE (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_18_WARMELTmiNORMAL_LAMBDA_cl1 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_19_WARMELTmiNORMAL_WRAP_NORMAL_LET1 (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_20_WARMELTmiNORMAL_LAMBDA_cl2 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_21_WARMELTmiNORMAL_WRAP_NORMAL_LETSEQ (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_22_WARMELTmiNORMAL_LAMBDA_cl3 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_23_WARMELTmiNORMAL_CHECK_CTYPE_NARGS (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_24_WARMELTmiNORMAL_LAMBDA_cl4 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_25_WARMELTmiNORMAL_NORMBIND_FAILANY (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_26_WARMELTmiNORMAL_NORMBIND_ANYBIND (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_27_WARMELTmiNORMAL_NORMBIND_FORMALBIND (meltclosure_ptr_t meltclosp_,
						 melt_ptr_t meltfirstargp_,
						 const melt_argdescr_cell_t
						 meltxargdescr_[],
						 union meltparam_un
						 *meltxargtab_,
						 const melt_argdescr_cell_t
						 meltxresdescr_[],
						 union meltparam_un
						 *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_28_WARMELTmiNORMAL_NORMBIND_LETBIND (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_29_WARMELTmiNORMAL_NORMBIND_FIXBIND (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_30_WARMELTmiNORMAL_NORMBIND_DEFINEDVALBIND (meltclosure_ptr_t
						     meltclosp_,
						     melt_ptr_t
						     meltfirstargp_,
						     const
						     melt_argdescr_cell_t
						     meltxargdescr_[],
						     union meltparam_un
						     *meltxargtab_,
						     const
						     melt_argdescr_cell_t
						     meltxresdescr_[],
						     union meltparam_un
						     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_31_WARMELTmiNORMAL_NORMBIND_CONSTRUCTBIND (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_32_WARMELTmiNORMAL_NORMEXP_SYMBOL (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_33_WARMELTmiNORMAL_GECTYP_SYMOCC (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_34_WARMELTmiNORMAL_GECTYP_MODVAROCC (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_35_WARMELTmiNORMAL_NORMEXP_CLASS (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_36_WARMELTmiNORMAL_NORMEXP_PRIMITIVE (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_37_WARMELTmiNORMAL_NORMEXP_HOOK_CALL (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_38_WARMELTmiNORMAL_NORMEXP_BOX (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_39_WARMELTmiNORMAL_NORMEXP_CONSTBOX (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_40_WARMELTmiNORMAL_NORMEXP_UNBOX (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_41_WARMELTmiNORMAL_NORMEXP_CODE_CHUNK (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_42_WARMELTmiNORMAL_LAMBDA_cl5 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_43_WARMELTmiNORMAL_NORMEXP_EXPR_CHUNK (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_44_WARMELTmiNORMAL_NORMEXP_CMATCHEXPR (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_45_WARMELTmiNORMAL_LAMBDA_cl6 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_46_WARMELTmiNORMAL_LAMBDA_cl7 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_47_WARMELTmiNORMAL_NORMEXP_FUNMATCHEXPR (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_48_WARMELTmiNORMAL_NORMEXP_APPLY (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_49_WARMELTmiNORMAL_NORMEXP_MSEND (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_50_WARMELTmiNORMAL_NORMEXP_RETURN (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_51_WARMELTmiNORMAL_NORMEXP_IF (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_52_WARMELTmiNORMAL_LAMBDA_cl8 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_53_WARMELTmiNORMAL_NORMEXP_IFELSE (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_54_WARMELTmiNORMAL_LAMBDA_cl9 (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_55_WARMELTmiNORMAL_LAMBDA_cl10 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_56_WARMELTmiNORMAL_NORMEXP_CPPIF (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_57_WARMELTmiNORMAL_LAMBDA_cl11 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_58_WARMELTmiNORMAL_LAMBDA_cl12 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_59_WARMELTmiNORMAL_NORMEXP_OR (meltclosure_ptr_t meltclosp_,
					melt_ptr_t meltfirstargp_,
					const melt_argdescr_cell_t
					meltxargdescr_[],
					union meltparam_un *meltxargtab_,
					const melt_argdescr_cell_t
					meltxresdescr_[],
					union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_60_WARMELTmiNORMAL_LAMBDA_cl13 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_61_WARMELTmiNORMAL_LAMBDA_cl14 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_62_WARMELTmiNORMAL_NORMEXP_PROGN (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_63_WARMELTmiNORMAL_LAMBDA_cl15 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_64_WARMELTmiNORMAL_NORMEXP_LET (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_65_WARMELTmiNORMAL_LAMBDA_cl16 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_66_WARMELTmiNORMAL_LAMBDA_cl17 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_67_WARMELTmiNORMAL_NORMEXP_UNSAFE_GET_FIELD (meltclosure_ptr_t
						      meltclosp_,
						      melt_ptr_t
						      meltfirstargp_,
						      const
						      melt_argdescr_cell_t
						      meltxargdescr_[],
						      union meltparam_un
						      *meltxargtab_,
						      const
						      melt_argdescr_cell_t
						      meltxresdescr_[],
						      union meltparam_un
						      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_68_WARMELTmiNORMAL_NORMEXP_GET_FIELD (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_69_WARMELTmiNORMAL_NORMEXP_UNSAFE_PUT_FIELDS (meltclosure_ptr_t
						       meltclosp_,
						       melt_ptr_t
						       meltfirstargp_,
						       const
						       melt_argdescr_cell_t
						       meltxargdescr_[],
						       union meltparam_un
						       *meltxargtab_,
						       const
						       melt_argdescr_cell_t
						       meltxresdescr_[],
						       union meltparam_un
						       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_70_WARMELTmiNORMAL_LAMBDA_cl18 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_71_WARMELTmiNORMAL_LAMBDA_cl19 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_72_WARMELTmiNORMAL_NORMEXP_PUT_FIELDS (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_73_WARMELTmiNORMAL_NORMEXP_SETQ (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_74_WARMELTmiNORMAL_LAMBDA_cl20 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_75_WARMELTmiNORMAL_NORMEXP_INSTANCE (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_76_WARMELTmiNORMAL_LAMBDA_cl21 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_77_WARMELTmiNORMAL_LAMBDA_cl22 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_78_WARMELTmiNORMAL_NORMEXP_FOREVER (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_79_WARMELTmiNORMAL_NORMEXP_EXIT (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_80_WARMELTmiNORMAL_NORMEXP_AGAIN (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_81_WARMELTmiNORMAL_NORMEXP_IFVARIADIC (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_82_WARMELTmiNORMAL_NORMEXP_COMPILEWARNING (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_83_WARMELTmiNORMAL_LAMBDA_cl23 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_84_WARMELTmiNORMAL_REPLACE_LAST_BY_RETURN (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_85_WARMELTmiNORMAL_LAMBDA_cl24 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_86_WARMELTmiNORMAL_NORMEXP_DEFUN (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_87_WARMELTmiNORMAL_LAMBDA_cl25 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_88_WARMELTmiNORMAL_LAMBDA_cl26 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_89_WARMELTmiNORMAL_NORMEXP_DEFHOOK (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_90_WARMELTmiNORMAL_LAMBDA_cl27 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_91_WARMELTmiNORMAL_NORMALIZE_LAMBDA (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_92_WARMELTmiNORMAL_LAMBDA_cl28 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_93_WARMELTmiNORMAL_NORMEXP_LAMBDA (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_94_WARMELTmiNORMAL_NORMEXP_MULTICALL (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_95_WARMELTmiNORMAL_LAMBDA_cl29 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_96_WARMELTmiNORMAL_LAMBDA_cl30 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_97_WARMELTmiNORMAL_LAMBDA_cl31 (meltclosure_ptr_t meltclosp_,
					 melt_ptr_t meltfirstargp_,
					 const melt_argdescr_cell_t
					 meltxargdescr_[],
					 union meltparam_un *meltxargtab_,
					 const melt_argdescr_cell_t
					 meltxresdescr_[],
					 union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_98_WARMELTmiNORMAL_NORMEXP_TUPLE (meltclosure_ptr_t meltclosp_,
					   melt_ptr_t meltfirstargp_,
					   const melt_argdescr_cell_t
					   meltxargdescr_[],
					   union meltparam_un *meltxargtab_,
					   const melt_argdescr_cell_t
					   meltxresdescr_[],
					   union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA (meltclosure_ptr_t meltclosp_,
						 melt_ptr_t meltfirstargp_,
						 const melt_argdescr_cell_t
						 meltxargdescr_[],
						 union meltparam_un
						 *meltxargtab_,
						 const melt_argdescr_cell_t
						 meltxresdescr_[],
						 union meltparam_un
						 *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un
						*meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un
						*meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un *meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_121_WARMELTmiNORMAL_NORMEXP_COMMENT (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_122_WARMELTmiNORMAL_NORMEXP_KEYWORD (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_123_WARMELTmiNORMAL_FILL_NORMAL_FORMALBIND (meltclosure_ptr_t
						     meltclosp_,
						     melt_ptr_t
						     meltfirstargp_,
						     const
						     melt_argdescr_cell_t
						     meltxargdescr_[],
						     union meltparam_un
						     *meltxargtab_,
						     const
						     melt_argdescr_cell_t
						     meltxresdescr_[],
						     union meltparam_un
						     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_124_WARMELTmiNORMAL_FILL_NORMAL_FORMALS (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_125_WARMELTmiNORMAL_FILL_NORMAL_EXPANSION (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_126_WARMELTmiNORMAL_LAMBDA_cl34 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_127_WARMELTmiNORMAL_NORMEXP_DEFPRIMITIVE (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_128_WARMELTmiNORMAL_NORMEXP_DEFCITERATOR (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_129_WARMELTmiNORMAL_NORMEXP_CITERATION (meltclosure_ptr_t meltclosp_,
						 melt_ptr_t meltfirstargp_,
						 const melt_argdescr_cell_t
						 meltxargdescr_[],
						 union meltparam_un
						 *meltxargtab_,
						 const melt_argdescr_cell_t
						 meltxresdescr_[],
						 union meltparam_un
						 *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_130_WARMELTmiNORMAL_LAMBDA_cl35 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_131_WARMELTmiNORMAL_LAMBDA_cl36 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_132_WARMELTmiNORMAL_LAMBDA_cl37 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_133_WARMELTmiNORMAL_LAMBDA_cl38 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_134_WARMELTmiNORMAL_LAMBDA_cl39 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_135_WARMELTmiNORMAL_LAMBDA_cl40 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_136_WARMELTmiNORMAL_LAMBDA_cl41 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_137_WARMELTmiNORMAL_NORMEXP_DEFCMATCHER (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_138_WARMELTmiNORMAL_NORMEXP_DEFUNMATCHER (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_139_WARMELTmiNORMAL_NORMEXP_DEFCLASS (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un
					       *meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un
					       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_140_WARMELTmiNORMAL_NORMEXP_DEFINSTANCE (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_141_WARMELTmiNORMAL_LAMBDA_cl42 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_142_WARMELTmiNORMAL_NORMEXP_DEFVAR (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_143_WARMELTmiNORMAL_NORMEXP_DEFINE (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un
					     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_144_WARMELTmiNORMAL_LAMBDA_cl43 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_145_WARMELTmiNORMAL_NORMEXP_DEFSELECTOR (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un
						  *meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un
						  *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_146_WARMELTmiNORMAL_LAMBDA_cl44 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_147_WARMELTmiNORMAL_NORMAL_EXPORTED_VALUE (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_148_WARMELTmiNORMAL_NORMEXP_EXPORT_VALUES (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_149_WARMELTmiNORMAL_NORMEXP_EXPORT_SYNONYM (meltclosure_ptr_t
						     meltclosp_,
						     melt_ptr_t
						     meltfirstargp_,
						     const
						     melt_argdescr_cell_t
						     meltxargdescr_[],
						     union meltparam_un
						     *meltxargtab_,
						     const
						     melt_argdescr_cell_t
						     meltxresdescr_[],
						     union meltparam_un
						     *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_150_WARMELTmiNORMAL_NORMEXP_EXPORT_CLASS (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_151_WARMELTmiNORMAL_LAMBDA_cl45 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_152_WARMELTmiNORMAL_LAMBDA_cl46 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_153_WARMELTmiNORMAL_NORMAL_EXPORTED_MACRO (meltclosure_ptr_t
						    meltclosp_,
						    melt_ptr_t meltfirstargp_,
						    const melt_argdescr_cell_t
						    meltxargdescr_[],
						    union meltparam_un
						    *meltxargtab_,
						    const melt_argdescr_cell_t
						    meltxresdescr_[],
						    union meltparam_un
						    *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_154_WARMELTmiNORMAL_NORMEXP_EXPORT_MACRO (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un
						   *meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un
						   *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_155_WARMELTmiNORMAL_NORMAL_EXPORTED_PATMACRO (meltclosure_ptr_t
						       meltclosp_,
						       melt_ptr_t
						       meltfirstargp_,
						       const
						       melt_argdescr_cell_t
						       meltxargdescr_[],
						       union meltparam_un
						       *meltxargtab_,
						       const
						       melt_argdescr_cell_t
						       meltxresdescr_[],
						       union meltparam_un
						       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_156_WARMELTmiNORMAL_NORMEXP_EXPORT_PATMACRO (meltclosure_ptr_t
						      meltclosp_,
						      melt_ptr_t
						      meltfirstargp_,
						      const
						      melt_argdescr_cell_t
						      meltxargdescr_[],
						      union meltparam_un
						      *meltxargtab_,
						      const
						      melt_argdescr_cell_t
						      meltxresdescr_[],
						      union meltparam_un
						      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_157_WARMELTmiNORMAL_NORMEXP_CURRENT_MODULE_ENVIRONMENT_REFERENCE
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_158_WARMELTmiNORMAL_NORMEXP_PARENT_MODULE_ENVIRONMENT
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_159_WARMELTmiNORMAL_NORMEXP_UPDATE_CURRENT_MODULE_ENVIRONMENT_REFERENCE
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_160_WARMELTmiNORMAL_NORMEXP_FETCH_PREDEFINED (meltclosure_ptr_t
						       meltclosp_,
						       melt_ptr_t
						       meltfirstargp_,
						       const
						       melt_argdescr_cell_t
						       meltxargdescr_[],
						       union meltparam_un
						       *meltxargtab_,
						       const
						       melt_argdescr_cell_t
						       meltxresdescr_[],
						       union meltparam_un
						       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_161_WARMELTmiNORMAL_LAMBDA_cl47 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_162_WARMELTmiNORMAL_NORMEXP_STORE_PREDEFINED (meltclosure_ptr_t
						       meltclosp_,
						       melt_ptr_t
						       meltfirstargp_,
						       const
						       melt_argdescr_cell_t
						       meltxargdescr_[],
						       union meltparam_un
						       *meltxargtab_,
						       const
						       melt_argdescr_cell_t
						       meltxresdescr_[],
						       union meltparam_un
						       *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_163_WARMELTmiNORMAL_LAMBDA_cl48 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_164_WARMELTmiNORMAL_NORMEXP_CHEADER (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un
					      *meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un
					      *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_165_WARMELTmiNORMAL_NORMEXP_CIMPLEMENT (meltclosure_ptr_t meltclosp_,
						 melt_ptr_t meltfirstargp_,
						 const melt_argdescr_cell_t
						 meltxargdescr_[],
						 union meltparam_un
						 *meltxargtab_,
						 const melt_argdescr_cell_t
						 meltxresdescr_[],
						 union meltparam_un
						 *meltxrestab_);




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_166_WARMELTmiNORMAL_NORMEXP_USE_PACKAGE_FROM_PKG_CONFIG
(meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_, const melt_argdescr_cell_t meltxargdescr_[],
union meltparam_un *meltxargtab_, const melt_argdescr_cell_t meltxresdescr_[],
union meltparam_un *meltxrestab_);



MELT_EXTERN void *melt_start_this_module (void *);
struct melt_initial_frame_st;

/*declare opaque initial frame: */
typedef struct melt_initial_frame_st meltinitial_frame_t;


/* define different names when debugging or not */
#if MELT_HAVE_DEBUG
MELT_EXTERN const char meltmodule_WARMELTmiNORMAL__melt_have_debug_enabled[];
#define melt_have_debug_string meltmodule_WARMELTmiNORMAL__melt_have_debug_enabled
#else /*!MELT_HAVE_DEBUG */
MELT_EXTERN const char meltmodule_WARMELTmiNORMAL__melt_have_debug_disabled[];
#define melt_have_debug_string meltmodule_WARMELTmiNORMAL__melt_have_debug_disabled
#endif /*!MELT_HAVE_DEBUG */



void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_0 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_1 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_2 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_3 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_4 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_5 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_6 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_7 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_8 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_9 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_10 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_11 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_12 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_13 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_14 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_15 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_16 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_17 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_18 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_19 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_20 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_21 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_22 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_23 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_24 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_25 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_26 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_27 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_28 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_29 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_30 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_31 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_32 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_33 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_34 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_35 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_36 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_37 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_38 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_39 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_40 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_41 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_42 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_43 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_44 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_45 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_46 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_47 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_48 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_49 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_50 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_51 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_52 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_53 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_54 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_55 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__initialmeltchunk_56 (meltinitial_frame_t *, char *);
void MELT_MODULE_VISIBILITY
meltmod__WARMELTmiNORMAL__forward_or_mark_module_start_frame (struct
							      melt_callframe_st
							      *fp,
							      int marking);



/**** warmelt-normal+04.c implementations ****/




melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un *meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un *meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST_melt =
    melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 94
    melt_ptr_t mcfr_varptr[94];
#define MELTFRAM_NBVARNUM 36
    long mcfr_varnum[36];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST_st *)
	meltfirstargp_;
      /* use arguments meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 94; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST nbval 94*/
  meltfram__.mcfr_nbvar = 94 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMEXP_LIST", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5521:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.PSLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.PSLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5522:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L1*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_LIST */
					    meltfrout->tabval[0])));;
      MELT_LOCATION ("warmelt-normal.melt:5522:/ cond");
      /*cond */ if ( /*_#IS_A__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V7*/ meltfptr[6] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5522:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check list recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5522) ? (5522) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V7*/ meltfptr[6] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V6*/ meltfptr[5] = /*_._IFELSE___V7*/ meltfptr[6];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5522:/ clear");
	     /*clear *//*_#IS_A__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V7*/ meltfptr[6] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V6*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5523:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L2*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5523:/ cond");
      /*cond */ if ( /*_#IS_A__L2*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5523:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check env"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5523) ? (5523) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5523:/ clear");
	     /*clear *//*_#IS_A__L2*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5524:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:5524:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[10] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5524:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5524) ? (5524) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[8] = /*_._IFELSE___V11*/ meltfptr[10];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5524:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[8] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5525:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5525:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[4] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5525:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[4];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5525;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_list recv";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V13*/ meltfptr[12] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V12*/ meltfptr[10] =
	      /*_.MELT_DEBUG_FUN__V13*/ meltfptr[12];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5525:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[4] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V13*/ meltfptr[12] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V12*/ meltfptr[10] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5525:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V12*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5526:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.LOCA_LOCATION__V15*/ meltfptr[10] = slot;
    };
    ;
    /*^cond */
    /*cond */ if ( /*_.LOCA_LOCATION__V15*/ meltfptr[10])	/*then */
      {
	/*^cond.then */
	/*_.SLOC__V16*/ meltfptr[15] = /*_.LOCA_LOCATION__V15*/ meltfptr[10];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5526:/ cond.else");

	/*_.SLOC__V16*/ meltfptr[15] = /*_.PSLOC__V5*/ meltfptr[4];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5527:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
  /*_.SARGS__V17*/ meltfptr[16] = slot;
    };
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5528:/ apply");
    /*apply */
    {
      /*_.CSYMREC__V18*/ meltfptr[17] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!CLONE_SYMBOL */ meltfrout->tabval[4])),
		    (melt_ptr_t) (( /*!konst_5_LISTREC_ */ meltfrout->
				   tabval[5])), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5529:/ apply");
    /*apply */
    {
      /*_.CSYMLIST__V19*/ meltfptr[18] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!CLONE_SYMBOL */ meltfrout->tabval[4])),
		    (melt_ptr_t) (( /*!konst_6_LIST_ */ meltfrout->
				   tabval[6])), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5530:/ apply");
    /*apply */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^apply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V16*/ meltfptr[15];
      /*^apply.arg */
      argtab[2].meltbp_cstring = "discr_pair";
      /*_.NORMDISCRPAIR__V20*/ meltfptr[19] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMAL_PREDEF */ meltfrout->tabval[7])),
		    (melt_ptr_t) (( /*!DISCR_PAIR */ meltfrout->tabval[8])),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING ""),
		    argtab, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5531:/ apply");
    /*apply */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^apply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V16*/ meltfptr[15];
      /*^apply.arg */
      argtab[2].meltbp_cstring = "discr_list";
      /*_.NORMDISCRLIST__V21*/ meltfptr[20] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMAL_PREDEF */ meltfrout->tabval[7])),
		    (melt_ptr_t) (( /*!DISCR_PAIR */ meltfrout->tabval[8])),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING ""),
		    argtab, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5532:/ apply");
    /*apply */
    {
      /*_.NEWENV__V22*/ meltfptr[21] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!FRESH_ENV */ meltfrout->tabval[9])),
		    (melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;
 /*_#NBARGS__L6*/ meltfnum[4] =
      (melt_multiple_length ((melt_ptr_t) ( /*_.SARGS__V17*/ meltfptr[16])));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5534:/ msend");
    /*msend */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^ojbmsend.arg */
      argtab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.CSYMLIST__V19*/ meltfptr[18];
      /*^ojbmsend.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^ojbmsend.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V16*/ meltfptr[15];
      /*_.PREPARE_CONSTRUCTOR_BINDING__V23*/ meltfptr[22] =
	meltgc_send ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
		     (melt_ptr_t) (( /*!PREPARE_CONSTRUCTOR_BINDING */
				    meltfrout->tabval[11])),
		     (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		     argtab, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    /*^apply */
    /*apply */
    {
      /*_.CONSBINDTUP__V24*/ meltfptr[23] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_TO_MULTIPLE */ meltfrout->tabval[10])),
		    (melt_ptr_t) ( /*_.PREPARE_CONSTRUCTOR_BINDING__V23*/
				  meltfptr[22]), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;
 /*_.RESBINDINGS__V25*/ meltfptr[24] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[12]))));;
    /*^compute */
 /*_.CONSLISTBIND__V26*/ meltfptr[25] =
      (melt_multiple_nth
       ((melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]), (-1)));;
    MELT_LOCATION ("warmelt-normal.melt:5539:/ quasiblock");


    /*^newclosure */
		 /*newclosure *//*_.LAMBDA___V28*/ meltfptr[27] =
      (melt_ptr_t)
      meltgc_new_closure ((meltobject_ptr_t)
			  (((melt_ptr_t) (MELT_PREDEF (DISCR_CLOSURE)))),
			  (meltroutine_ptr_t) (( /*!konst_19 */ meltfrout->
						tabval[19])), (1));
    ;
    /*^putclosedv */
    /*putclosv */
    melt_assertmsg ("putclosv checkclo",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.LAMBDA___V28*/ meltfptr[27])) ==
		    MELTOBMAG_CLOSURE);
    melt_assertmsg ("putclosv checkoff", 0 >= 0
		    && 0 <
		    melt_closure_size ((melt_ptr_t)
				       ( /*_.LAMBDA___V28*/ meltfptr[27])));
    ((meltclosure_ptr_t) /*_.LAMBDA___V28*/ meltfptr[27])->tabval[0] =
      (melt_ptr_t) ( /*_.SLOC__V16*/ meltfptr[15]);
    ;
    /*_.LAMBDA___V27*/ meltfptr[26] = /*_.LAMBDA___V28*/ meltfptr[27];;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5537:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.LAMBDA___V27*/ meltfptr[26];
      /*_.TUPRECLOC__V29*/ meltfptr[28] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!MULTIPLE_MAP */ meltfrout->tabval[13])),
		    (melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5550:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L7*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5550:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L7*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L8*/ meltfnum[7] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5550:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[11];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L8*/ meltfnum[7];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5550;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_list csymrec=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CSYMREC__V18*/ meltfptr[17];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " consbindtup=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSBINDTUP__V24*/ meltfptr[23];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " conslistbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSLISTBIND__V26*/ meltfptr[25];
	      /*^apply.arg */
	      argtab[9].meltbp_cstring = " tuprecloc=";
	      /*^apply.arg */
	      argtab[10].meltbp_aptr =
		(melt_ptr_t *) & /*_.TUPRECLOC__V29*/ meltfptr[28];
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[30] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V30*/ meltfptr[29] =
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[30];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5550:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L8*/ meltfnum[7] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V31*/ meltfptr[30] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V30*/ meltfptr[29] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5550:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L7*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V30*/ meltfptr[29] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5553:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L9*/ meltfnum[7] =
	melt_is_instance_of ((melt_ptr_t)
			     ( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LIST_BINDING */ meltfrout->tabval[20])));;
      MELT_LOCATION ("warmelt-normal.melt:5553:/ cond");
      /*cond */ if ( /*_#IS_A__L9*/ meltfnum[7])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V33*/ meltfptr[29] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5553:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check conslistbind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5553) ? (5553) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V33*/ meltfptr[29] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V32*/ meltfptr[30] = /*_._IFELSE___V33*/ meltfptr[29];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5553:/ clear");
	     /*clear *//*_#IS_A__L9*/ meltfnum[7] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V33*/ meltfptr[29] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V32*/ meltfptr[30] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5554:/ quasiblock");


    /*^multiapply */
    /*multiapply 4args, 1x.res */
    {
      union meltparam_un argtab[3];

      union meltparam_un restab[1];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V16*/ meltfptr[15];
      /*^multiapply.xres */
      restab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.NBINDINGS__V36*/ meltfptr[35];
      /*^multiapply.appl */
      /*_.NARGS__V35*/ meltfptr[34] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_TUPLE */ meltfrout->tabval[21])),
		    (melt_ptr_t) ( /*_.SARGS__V17*/ meltfptr[16]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		    argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5557:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L10*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5557:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L11*/ meltfnum[7] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5557:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L11*/ meltfnum[7];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5557;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_list nargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NARGS__V35*/ meltfptr[34];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nbindings=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBINDINGS__V36*/ meltfptr[35];
	      /*_.MELT_DEBUG_FUN__V38*/ meltfptr[37] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V37*/ meltfptr[36] =
	      /*_.MELT_DEBUG_FUN__V38*/ meltfptr[37];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5557:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[7] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V38*/ meltfptr[37] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V37*/ meltfptr[36] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5557:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V37*/ meltfptr[36] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5558:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L12*/ meltfnum[7] =
	(( /*_.NBINDINGS__V36*/ meltfptr[35]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBINDINGS__V36*/ meltfptr[35])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:5558:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L12*/ meltfnum[7])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V40*/ meltfptr[36] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5558:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nbindings"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5558) ? (5558) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V40*/ meltfptr[36] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V39*/ meltfptr[37] = /*_._IFELSE___V40*/ meltfptr[36];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5558:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L12*/ meltfnum[7] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V40*/ meltfptr[36] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V39*/ meltfptr[37] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5559:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.NBINDINGS__V36*/ meltfptr[35];
      /*_.LIST_APPEND2LIST__V41*/ meltfptr[36] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_APPEND2LIST */ meltfrout->tabval[22])),
		    (melt_ptr_t) ( /*_.RESBINDINGS__V25*/ meltfptr[24]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;
    /*citerblock FOREACH_IN_LIST */
    {
      /* start foreach_in_list meltcit1__EACHLIST */
      for ( /*_.CURPAIRBIND__V42*/ meltfptr[41] =
	   melt_list_first ((melt_ptr_t) /*_.RESBINDINGS__V25*/ meltfptr[24]);
	   melt_magic_discr ((melt_ptr_t) /*_.CURPAIRBIND__V42*/ meltfptr[41])
	   == MELTOBMAG_PAIR;
	   /*_.CURPAIRBIND__V42*/ meltfptr[41] =
	   melt_pair_tail ((melt_ptr_t) /*_.CURPAIRBIND__V42*/ meltfptr[41]))
	{
	  /*_.CURBIND__V43*/ meltfptr[42] =
	    melt_pair_head ((melt_ptr_t) /*_.CURPAIRBIND__V42*/ meltfptr[41]);



	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5563:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.CURBIND__V43*/ meltfptr[42];
	    /*_.PUT_ENV__V44*/ meltfptr[43] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!PUT_ENV */ meltfrout->tabval[23])),
			  (melt_ptr_t) ( /*_.NEWENV__V22*/ meltfptr[21]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;
	}			/* end foreach_in_list meltcit1__EACHLIST */
     /*_.CURPAIRBIND__V42*/ meltfptr[41] = NULL;
     /*_.CURBIND__V43*/ meltfptr[42] = NULL;


      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5560:/ clear");
	    /*clear *//*_.CURPAIRBIND__V42*/ meltfptr[41] = 0;
      /*^clear */
	    /*clear *//*_.CURBIND__V43*/ meltfptr[42] = 0;
      /*^clear */
	    /*clear *//*_.PUT_ENV__V44*/ meltfptr[43] = 0;
    }				/*endciterblock FOREACH_IN_LIST */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit2__EACHTUP */
      long meltcit2__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NARGS__V35*/ meltfptr[34]);
      for ( /*_#NIX__L13*/ meltfnum[0] = 0;
	   ( /*_#NIX__L13*/ meltfnum[0] >= 0)
	   && ( /*_#NIX__L13*/ meltfnum[0] < meltcit2__EACHTUP_ln);
	/*_#NIX__L13*/ meltfnum[0]++)
	{
	  /*_.CURNARG__V45*/ meltfptr[44] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.NARGS__V35*/ meltfptr[34]),
			       /*_#NIX__L13*/ meltfnum[0]);




	  MELT_CHECK_SIGNAL ();
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5568:/ msend");
	  /*msend */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NEWENV__V22*/ meltfptr[21];
	    /*_.GET_CTYPE__V46*/ meltfptr[45] =
	      meltgc_send ((melt_ptr_t) ( /*_.CURNARG__V45*/ meltfptr[44]),
			   (melt_ptr_t) (( /*!GET_CTYPE */ meltfrout->
					  tabval[24])), (MELTBPARSTR_PTR ""),
			   argtab, "", (union meltparam_un *) 0);
	  }
	  ;
  /*_#exeq__L14*/ meltfnum[7] =
	    (( /*_.GET_CTYPE__V46*/ meltfptr[45]) !=
	     (( /*!CTYPE_VALUE */ meltfrout->tabval[25])));;
	  MELT_LOCATION ("warmelt-normal.melt:5568:/ cond");
	  /*cond */ if ( /*_#exeq__L14*/ meltfnum[7])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {


#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:5570:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#MELT_NEED_DBG__L15*/ meltfnum[14] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:5570:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L15*/ meltfnum[14])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	/*_#MELT_CALLCOUNT__L16*/ meltfnum[15] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:5570:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[5];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L16*/ meltfnum[15];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 5570;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "normexp_list bad curnarg";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CURNARG__V45*/ meltfptr[44];
			  /*_.MELT_DEBUG_FUN__V49*/ meltfptr[48] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[3])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V48*/ meltfptr[47] =
			  /*_.MELT_DEBUG_FUN__V49*/ meltfptr[48];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:5570:/ clear");
		  /*clear *//*_#MELT_CALLCOUNT__L16*/ meltfnum[15] = 0;
			/*^clear */
		  /*clear *//*_.MELT_DEBUG_FUN__V49*/ meltfptr[48] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

       /*_._IF___V48*/ meltfptr[47] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:5570:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		/*clear *//*_#MELT_NEED_DBG__L15*/ meltfnum[14] = 0;
		  /*^clear */
		/*clear *//*_._IF___V48*/ meltfptr[47] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

		{
		  MELT_LOCATION ("warmelt-normal.melt:5571:/ locexp");
		  /* error_plain */
		    melt_error_str ((melt_ptr_t)
				    ( /*_.SLOC__V16*/ meltfptr[15]),
				    ("(LIST ...) arguments should be values"),
				    (melt_ptr_t) 0);
		}
		;
		MELT_LOCATION ("warmelt-normal.melt:5572:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = NULL;;

		{
		  MELT_LOCATION ("warmelt-normal.melt:5572:/ locexp");
		  /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
		  if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
		    melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
		  ;
		}
		;
		/*^finalreturn */
		;
		/*finalret */ goto meltlabend_rout;
		MELT_LOCATION ("warmelt-normal.melt:5569:/ quasiblock");


		/*_.PROGN___V51*/ meltfptr[47] =
		  /*_.RETURN___V50*/ meltfptr[48];;
		/*^compute */
		/*_._IF___V47*/ meltfptr[46] =
		  /*_.PROGN___V51*/ meltfptr[47];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5568:/ clear");
	      /*clear *//*_.RETURN___V50*/ meltfptr[48] = 0;
		/*^clear */
	      /*clear *//*_.PROGN___V51*/ meltfptr[47] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_._IF___V47*/ meltfptr[46] = NULL;;
	    }
	  ;
	  if ( /*_#NIX__L13*/ meltfnum[0] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit2__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5565:/ clear");
	    /*clear *//*_.CURNARG__V45*/ meltfptr[44] = 0;
      /*^clear */
	    /*clear *//*_#NIX__L13*/ meltfnum[0] = 0;
      /*^clear */
	    /*clear *//*_.GET_CTYPE__V46*/ meltfptr[45] = 0;
      /*^clear */
	    /*clear *//*_#exeq__L14*/ meltfnum[7] = 0;
      /*^clear */
	    /*clear *//*_._IF___V47*/ meltfptr[46] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5573:/ blockmultialloc");
    /*multiallocblock */
    {
      struct meltletrec_1_st
      {
	struct MELT_OBJECT_STRUCT (4) rinst_0__CRECLOCC;
	struct MELT_MULTIPLE_STRUCT (1) rtup_1__NBDY;
	struct MELT_OBJECT_STRUCT (6) rinst_2__NLETREC;
	struct MELT_OBJECT_STRUCT (4) rinst_3__CLISTBIND;
	struct MELT_OBJECT_STRUCT (4) rinst_4__CLISTLOCC;
	struct meltpair_st rpair_5__CLISTBIND_x1;
	struct meltlist_st rlist_6__CLISTBINDLIST;
	long meltletrec_1_endgap;
      } *meltletrec_1_ptr = 0;
      meltletrec_1_ptr =
	(struct meltletrec_1_st *)
	meltgc_allocate (sizeof (struct meltletrec_1_st), 0);
      /*^blockmultialloc.initfill */

      /*iniobj rinst_0__CRECLOCC */
  /*_.CRECLOCC__V53*/ meltfptr[47] =
	(melt_ptr_t) & meltletrec_1_ptr->rinst_0__CRECLOCC;
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr isobj rinst_0__CRECLOCC",
			melt_magic_discr ((melt_ptr_t)
					  (( /*!CLASS_NREP_LOCSYMOCC */
					    meltfrout->tabval[26]))) ==
			MELTOBMAG_OBJECT);
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr objmagic rinst_0__CRECLOCC",
			((meltobject_ptr_t)
			 (( /*!CLASS_NREP_LOCSYMOCC */ meltfrout->
			   tabval[26])))->meltobj_magic == MELTOBMAG_OBJECT);
      meltletrec_1_ptr->rinst_0__CRECLOCC.meltobj_class =
	(meltobject_ptr_t) (( /*!CLASS_NREP_LOCSYMOCC */ meltfrout->
			     tabval[26]));
      meltletrec_1_ptr->rinst_0__CRECLOCC.obj_hash = melt_nonzerohash ();
      meltletrec_1_ptr->rinst_0__CRECLOCC.obj_len = 4;

      /*inimult rtup_1__NBDY */
 /*_.NBDY__V54*/ meltfptr[53] =
	(melt_ptr_t) & meltletrec_1_ptr->rtup_1__NBDY;
      meltletrec_1_ptr->rtup_1__NBDY.discr =
	(meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_MULTIPLE))));
      meltletrec_1_ptr->rtup_1__NBDY.nbval = 1;

      /*iniobj rinst_2__NLETREC */
  /*_.NLETREC__V55*/ meltfptr[54] =
	(melt_ptr_t) & meltletrec_1_ptr->rinst_2__NLETREC;
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr isobj rinst_2__NLETREC",
			melt_magic_discr ((melt_ptr_t)
					  (( /*!CLASS_NREP_LETREC */
					    meltfrout->tabval[27]))) ==
			MELTOBMAG_OBJECT);
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr objmagic rinst_2__NLETREC",
			((meltobject_ptr_t)
			 (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[27])))->
			meltobj_magic == MELTOBMAG_OBJECT);
      meltletrec_1_ptr->rinst_2__NLETREC.meltobj_class =
	(meltobject_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[27]));
      meltletrec_1_ptr->rinst_2__NLETREC.obj_hash = melt_nonzerohash ();
      meltletrec_1_ptr->rinst_2__NLETREC.obj_len = 6;


      /*iniobj rinst_3__CLISTBIND */
  /*_.CLISTBIND__V56*/ meltfptr[55] =
	(melt_ptr_t) & meltletrec_1_ptr->rinst_3__CLISTBIND;
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr isobj rinst_3__CLISTBIND",
			melt_magic_discr ((melt_ptr_t)
					  (( /*!CLASS_NORMAL_LET_BINDING */
					    meltfrout->tabval[28]))) ==
			MELTOBMAG_OBJECT);
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr objmagic rinst_3__CLISTBIND",
			((meltobject_ptr_t)
			 (( /*!CLASS_NORMAL_LET_BINDING */ meltfrout->
			   tabval[28])))->meltobj_magic == MELTOBMAG_OBJECT);
      meltletrec_1_ptr->rinst_3__CLISTBIND.meltobj_class =
	(meltobject_ptr_t) (( /*!CLASS_NORMAL_LET_BINDING */ meltfrout->
			     tabval[28]));
      meltletrec_1_ptr->rinst_3__CLISTBIND.obj_hash = melt_nonzerohash ();
      meltletrec_1_ptr->rinst_3__CLISTBIND.obj_len = 4;


      /*iniobj rinst_4__CLISTLOCC */
  /*_.CLISTLOCC__V57*/ meltfptr[56] =
	(melt_ptr_t) & meltletrec_1_ptr->rinst_4__CLISTLOCC;
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr isobj rinst_4__CLISTLOCC",
			melt_magic_discr ((melt_ptr_t)
					  (( /*!CLASS_NREP_LOCSYMOCC */
					    meltfrout->tabval[26]))) ==
			MELTOBMAG_OBJECT);
      if (MELT_LIKELY (!melt_prohibit_garbcoll))
	melt_assertmsg ("iniobj check.discr objmagic rinst_4__CLISTLOCC",
			((meltobject_ptr_t)
			 (( /*!CLASS_NREP_LOCSYMOCC */ meltfrout->
			   tabval[26])))->meltobj_magic == MELTOBMAG_OBJECT);
      meltletrec_1_ptr->rinst_4__CLISTLOCC.meltobj_class =
	(meltobject_ptr_t) (( /*!CLASS_NREP_LOCSYMOCC */ meltfrout->
			     tabval[26]));
      meltletrec_1_ptr->rinst_4__CLISTLOCC.obj_hash = melt_nonzerohash ();
      meltletrec_1_ptr->rinst_4__CLISTLOCC.obj_len = 4;

      /*inipair rpair_5__CLISTBIND_x1 */
   /*_.CLISTBIND__V58*/ meltfptr[57] =
	(melt_ptr_t) & meltletrec_1_ptr->rpair_5__CLISTBIND_x1;
      meltletrec_1_ptr->rpair_5__CLISTBIND_x1.discr =
	(meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_PAIR))));

      /*inilist rlist_6__CLISTBINDLIST */
   /*_.CLISTBINDLIST__V59*/ meltfptr[58] =
	(melt_ptr_t) & meltletrec_1_ptr->rlist_6__CLISTBINDLIST;
      meltletrec_1_ptr->rlist_6__CLISTBINDLIST.discr =
	(meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_LIST))));



      MELT_LOCATION ("warmelt-normal.melt:5574:/ putslot");
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CRECLOCC__V53*/ meltfptr[47]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CRECLOCC__V53*/ meltfptr[47]), (0),
			    ( /*_.SLOC__V16*/ meltfptr[15]), "NREP_LOC");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_SYMB",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CRECLOCC__V53*/ meltfptr[47]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CRECLOCC__V53*/ meltfptr[47]), (1),
			    ( /*_.CSYMREC__V18*/ meltfptr[17]), "NOCC_SYMB");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_CTYP",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CRECLOCC__V53*/ meltfptr[47]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CRECLOCC__V53*/ meltfptr[47]), (2),
			    (( /*!CTYPE_VALUE */ meltfrout->tabval[25])),
			    "NOCC_CTYP");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_BIND",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CRECLOCC__V53*/ meltfptr[47]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CRECLOCC__V53*/ meltfptr[47]), (3),
			    ( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
			    "NOCC_BIND");
      ;
      /*^touch */
      meltgc_touch ( /*_.CRECLOCC__V53*/ meltfptr[47]);
      ;
      MELT_LOCATION ("warmelt-normal.melt:5579:/ putuple");
      /*putupl#40 */
      melt_assertmsg ("putupl [:5579] #40 checktup",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NBDY__V54*/ meltfptr[53])) ==
		      MELTOBMAG_MULTIPLE);
      melt_assertmsg ("putupl [:5579] #40 checkoff",
		      (0 >= 0
		       && 0 <
		       melt_multiple_length ((melt_ptr_t)
					     ( /*_.NBDY__V54*/
					      meltfptr[53]))));
      ((meltmultiple_ptr_t) ( /*_.NBDY__V54*/ meltfptr[53]))->tabval[0] =
	(melt_ptr_t) ( /*_.CRECLOCC__V53*/ meltfptr[47]);
      ;
      /*^touch */
      meltgc_touch ( /*_.NBDY__V54*/ meltfptr[53]);
      ;
      MELT_LOCATION ("warmelt-normal.melt:5580:/ putslot");
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (0),
			    ( /*_.SLOC__V16*/ meltfptr[15]), "NREP_LOC");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NLET_BINDINGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (1),
			    ( /*_.CONSBINDTUP__V24*/ meltfptr[23]),
			    "NLET_BINDINGS");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NLET_BODY",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (2),
			    ( /*_.NBDY__V54*/ meltfptr[53]), "NLET_BODY");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NLETREC_FILL_BINDINGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (3),
			    (( /*nil */ NULL)), "NLETREC_FILL_BINDINGS");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NLETREC_BODY_BINDINGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (4),
			    (( /*nil */ NULL)), "NLETREC_BODY_BINDINGS");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NLETREC_LOCSYMS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.NLETREC__V55*/ meltfptr[54])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.NLETREC__V55*/ meltfptr[54]), (5),
			    ( /*_.TUPRECLOC__V29*/ meltfptr[28]),
			    "NLETREC_LOCSYMS");
      ;
      /*^touch */
      meltgc_touch ( /*_.NLETREC__V55*/ meltfptr[54]);
      ;
      MELT_LOCATION ("warmelt-normal.melt:5588:/ putslot");
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @BINDER",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBIND__V56*/ meltfptr[55]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTBIND__V56*/ meltfptr[55]), (0),
			    ( /*_.CSYMLIST__V19*/ meltfptr[18]), "BINDER");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LETBIND_TYPE",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBIND__V56*/ meltfptr[55]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTBIND__V56*/ meltfptr[55]), (1),
			    (( /*!CTYPE_VALUE */ meltfrout->tabval[25])),
			    "LETBIND_TYPE");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LETBIND_EXPR",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBIND__V56*/ meltfptr[55]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTBIND__V56*/ meltfptr[55]), (2),
			    ( /*_.NLETREC__V55*/ meltfptr[54]),
			    "LETBIND_EXPR");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LETBIND_LOC",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBIND__V56*/ meltfptr[55]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTBIND__V56*/ meltfptr[55]), (3),
			    ( /*_.SLOC__V16*/ meltfptr[15]), "LETBIND_LOC");
      ;
      /*^touch */
      meltgc_touch ( /*_.CLISTBIND__V56*/ meltfptr[55]);
      ;
      MELT_LOCATION ("warmelt-normal.melt:5593:/ putslot");
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTLOCC__V57*/ meltfptr[56]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTLOCC__V57*/ meltfptr[56]), (0),
			    ( /*_.SLOC__V16*/ meltfptr[15]), "NREP_LOC");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_SYMB",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTLOCC__V57*/ meltfptr[56]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTLOCC__V57*/ meltfptr[56]), (1),
			    ( /*_.CSYMLIST__V19*/ meltfptr[18]), "NOCC_SYMB");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_CTYP",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTLOCC__V57*/ meltfptr[56]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTLOCC__V57*/ meltfptr[56]), (2),
			    (( /*!CTYPE_VALUE */ meltfrout->tabval[25])),
			    "NOCC_CTYP");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @NOCC_BIND",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTLOCC__V57*/ meltfptr[56]))
		      == MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.CLISTLOCC__V57*/ meltfptr[56]), (3),
			    ( /*_.CLISTBIND__V56*/ meltfptr[55]),
			    "NOCC_BIND");
      ;
      /*^touch */
      meltgc_touch ( /*_.CLISTLOCC__V57*/ meltfptr[56]);
      ;
      MELT_LOCATION ("warmelt-normal.melt:5598:/ putpairhead");
      /*putpairhead */
      melt_assertmsg ("putpairhead /d checkpair",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBIND__V58*/ meltfptr[57]))
		      == MELTOBMAG_PAIR);
      ((meltpair_ptr_t) ( /*_.CLISTBIND__V58*/ meltfptr[57]))->hd =
	(melt_ptr_t) ( /*_.CLISTBIND__V56*/ meltfptr[55]);
      ;
      /*^touch */
      meltgc_touch ( /*_.CLISTBIND__V58*/ meltfptr[57]);
      ;
      /*^putlist */
      /*putlist */
      melt_assertmsg ("putlist checklist",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.CLISTBINDLIST__V59*/
					 meltfptr[58])) == MELTOBMAG_LIST);
      ((meltlist_ptr_t) ( /*_.CLISTBINDLIST__V59*/ meltfptr[58]))->first =
	(meltpair_ptr_t) ( /*_.CLISTBIND__V58*/ meltfptr[57]);
      ((meltlist_ptr_t) ( /*_.CLISTBINDLIST__V59*/ meltfptr[58]))->last =
	(meltpair_ptr_t) ( /*_.CLISTBIND__V58*/ meltfptr[57]);
      ;
      /*^touch */
      meltgc_touch ( /*_.CLISTBINDLIST__V59*/ meltfptr[58]);
      ;
      /*citerblock FOREACH_IN_MULTIPLE */
      {
	/* start foreach_in_multiple meltcit3__EACHTUP */
	long meltcit3__EACHTUP_ln =
	  melt_multiple_length ((melt_ptr_t) /*_.NARGS__V35*/ meltfptr[34]);
	for ( /*_#NIX__L17*/ meltfnum[15] = 0;
	     ( /*_#NIX__L17*/ meltfnum[15] >= 0)
	     && ( /*_#NIX__L17*/ meltfnum[15] < meltcit3__EACHTUP_ln);
	/*_#NIX__L17*/ meltfnum[15]++)
	  {
	    /*_.CURNARG__V60*/ meltfptr[59] =
	      melt_multiple_nth ((melt_ptr_t)
				 ( /*_.NARGS__V35*/ meltfptr[34]),
				 /*_#NIX__L17*/ meltfnum[15]);



	    MELT_LOCATION ("warmelt-normal.melt:5603:/ quasiblock");


   /*_.CURPAIRB__V61*/ meltfptr[60] =
	      (melt_multiple_nth
	       ((melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]),
		( /*_#NIX__L17*/ meltfnum[15])));;
	    /*^compute */
   /*_#plI__L18*/ meltfnum[14] =
	      (( /*_#NIX__L17*/ meltfnum[15]) + (1));;
	    /*^compute */
   /*_.NEXTB__V62*/ meltfptr[61] =
	      (melt_multiple_nth
	       ((melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]),
		( /*_#plI__L18*/ meltfnum[14])));;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:5606:/ locexp");
	      melt_put_int ((melt_ptr_t) ( /*_.CURPAIRB__V61*/ meltfptr[60]),
			    ( /*_#NIX__L17*/ meltfnum[15]));
	    }
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5607:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		melt_dbgcounter++;
#endif
		;
	      }
	      ;

	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#MELT_NEED_DBG__L19*/ meltfnum[18] =
		/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		;;
	      MELT_LOCATION ("warmelt-normal.melt:5607:/ cond");
	      /*cond */ if ( /*_#MELT_NEED_DBG__L19*/ meltfnum[18])	/*then */
		{
		  /*^cond.then */
		  /*^block */
		  /*anyblock */
		  {

       /*_#MELT_CALLCOUNT__L20*/ meltfnum[19] =
		      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		      meltcallcount	/* melt_callcount debugging */
#else
		      0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		      ;;

		    MELT_CHECK_SIGNAL ();
		    ;
		    MELT_LOCATION ("warmelt-normal.melt:5607:/ apply");
		    /*apply */
		    {
		      union meltparam_un argtab[5];
		      memset (&argtab, 0, sizeof (argtab));
		      /*^apply.arg */
		      argtab[0].meltbp_long =
			/*_#MELT_CALLCOUNT__L20*/ meltfnum[19];
		      /*^apply.arg */
		      argtab[1].meltbp_cstring = "warmelt-normal.melt";
		      /*^apply.arg */
		      argtab[2].meltbp_long = 5607;
		      /*^apply.arg */
		      argtab[3].meltbp_cstring = "normexp_list curpairb";
		      /*^apply.arg */
		      argtab[4].meltbp_aptr =
			(melt_ptr_t *) & /*_.CURPAIRB__V61*/ meltfptr[60];
		      /*_.MELT_DEBUG_FUN__V64*/ meltfptr[63] =
			melt_apply ((meltclosure_ptr_t)
				    (( /*!MELT_DEBUG_FUN */ meltfrout->
				      tabval[3])),
				    (melt_ptr_t) (( /*nil */ NULL)),
				    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_PTR ""), argtab, "",
				    (union meltparam_un *) 0);
		    }
		    ;
		    /*_._IF___V63*/ meltfptr[62] =
		      /*_.MELT_DEBUG_FUN__V64*/ meltfptr[63];;
		    /*epilog */

		    MELT_LOCATION ("warmelt-normal.melt:5607:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L20*/ meltfnum[19] = 0;
		    /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V64*/ meltfptr[63] = 0;
		  }
		  ;
		}
	      else
		{		/*^cond.else */

      /*_._IF___V63*/ meltfptr[62] = NULL;;
		}
	      ;

	      {
		MELT_LOCATION ("warmelt-normal.melt:5607:/ locexp");
		/*void */ (void) 0;
	      }
	      ;
	      /*^quasiblock */


	      /*epilog */

	      /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L19*/ meltfnum[18] = 0;
	      /*^clear */
	       /*clear *//*_._IF___V63*/ meltfptr[62] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*void */ (void) 0;
	      }
	      ;
	      /*epilog */
	    }

#endif /*MELT_HAVE_DEBUG */
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5608:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#IS_A__L21*/ meltfnum[19] =
		melt_is_instance_of ((melt_ptr_t)
				     ( /*_.CURPAIRB__V61*/ meltfptr[60]),
				     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[29])));;
	      MELT_LOCATION ("warmelt-normal.melt:5608:/ cond");
	      /*cond */ if ( /*_#IS_A__L21*/ meltfnum[19])	/*then */
		{
		  /*^cond.then */
		  /*_._IFELSE___V66*/ meltfptr[62] = ( /*nil */ NULL);;
		}
	      else
		{
		  MELT_LOCATION ("warmelt-normal.melt:5608:/ cond.else");

		  /*^block */
		  /*anyblock */
		  {




		    {
		      /*^locexp */
		      melt_assert_failed (("check curpairb"),
					  ("warmelt-normal.melt")
					  ? ("warmelt-normal.melt") :
					  __FILE__,
					  (5608) ? (5608) : __LINE__,
					  __FUNCTION__);
		      ;
		    }
		    ;
		 /*clear *//*_._IFELSE___V66*/ meltfptr[62] = 0;
		    /*epilog */
		  }
		  ;
		}
	      ;
	      /*^compute */
	      /*_.IFCPP___V65*/ meltfptr[63] =
		/*_._IFELSE___V66*/ meltfptr[62];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5608:/ clear");
	       /*clear *//*_#IS_A__L21*/ meltfnum[19] = 0;
	      /*^clear */
	       /*clear *//*_._IFELSE___V66*/ meltfptr[62] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*_.IFCPP___V65*/ meltfptr[63] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5609:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#GET_INT__L22*/ meltfnum[18] =
		(melt_get_int
		 ((melt_ptr_t) ( /*_.CURPAIRB__V61*/ meltfptr[60])));;
	      /*^compute */
     /*_#eqeqI__L23*/ meltfnum[19] =
		(( /*_#GET_INT__L22*/ meltfnum[18]) ==
		 ( /*_#NIX__L17*/ meltfnum[15]));;
	      MELT_LOCATION ("warmelt-normal.melt:5609:/ cond");
	      /*cond */ if ( /*_#eqeqI__L23*/ meltfnum[19])	/*then */
		{
		  /*^cond.then */
		  /*_._IFELSE___V68*/ meltfptr[67] = ( /*nil */ NULL);;
		}
	      else
		{
		  MELT_LOCATION ("warmelt-normal.melt:5609:/ cond.else");

		  /*^block */
		  /*anyblock */
		  {




		    {
		      /*^locexp */
		      melt_assert_failed (("check curpairb objnum"),
					  ("warmelt-normal.melt")
					  ? ("warmelt-normal.melt") :
					  __FILE__,
					  (5609) ? (5609) : __LINE__,
					  __FUNCTION__);
		      ;
		    }
		    ;
		 /*clear *//*_._IFELSE___V68*/ meltfptr[67] = 0;
		    /*epilog */
		  }
		  ;
		}
	      ;
	      /*^compute */
	      /*_.IFCPP___V67*/ meltfptr[62] =
		/*_._IFELSE___V68*/ meltfptr[67];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5609:/ clear");
	       /*clear *//*_#GET_INT__L22*/ meltfnum[18] = 0;
	      /*^clear */
	       /*clear *//*_#eqeqI__L23*/ meltfnum[19] = 0;
	      /*^clear */
	       /*clear *//*_._IFELSE___V68*/ meltfptr[67] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*_.IFCPP___V67*/ meltfptr[62] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5610:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		melt_dbgcounter++;
#endif
		;
	      }
	      ;

	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#MELT_NEED_DBG__L24*/ meltfnum[18] =
		/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		;;
	      MELT_LOCATION ("warmelt-normal.melt:5610:/ cond");
	      /*cond */ if ( /*_#MELT_NEED_DBG__L24*/ meltfnum[18])	/*then */
		{
		  /*^cond.then */
		  /*^block */
		  /*anyblock */
		  {

       /*_#MELT_CALLCOUNT__L25*/ meltfnum[19] =
		      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		      meltcallcount	/* melt_callcount debugging */
#else
		      0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		      ;;

		    MELT_CHECK_SIGNAL ();
		    ;
		    MELT_LOCATION ("warmelt-normal.melt:5610:/ apply");
		    /*apply */
		    {
		      union meltparam_un argtab[7];
		      memset (&argtab, 0, sizeof (argtab));
		      /*^apply.arg */
		      argtab[0].meltbp_long =
			/*_#MELT_CALLCOUNT__L25*/ meltfnum[19];
		      /*^apply.arg */
		      argtab[1].meltbp_cstring = "warmelt-normal.melt";
		      /*^apply.arg */
		      argtab[2].meltbp_long = 5610;
		      /*^apply.arg */
		      argtab[3].meltbp_cstring = "normexp_list curnarg=";
		      /*^apply.arg */
		      argtab[4].meltbp_aptr =
			(melt_ptr_t *) & /*_.CURNARG__V60*/ meltfptr[59];
		      /*^apply.arg */
		      argtab[5].meltbp_cstring = " nextb=";
		      /*^apply.arg */
		      argtab[6].meltbp_aptr =
			(melt_ptr_t *) & /*_.NEXTB__V62*/ meltfptr[61];
		      /*_.MELT_DEBUG_FUN__V70*/ meltfptr[69] =
			melt_apply ((meltclosure_ptr_t)
				    (( /*!MELT_DEBUG_FUN */ meltfrout->
				      tabval[3])),
				    (melt_ptr_t) (( /*nil */ NULL)),
				    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				     MELTBPARSTR_PTR ""), argtab, "",
				    (union meltparam_un *) 0);
		    }
		    ;
		    /*_._IF___V69*/ meltfptr[67] =
		      /*_.MELT_DEBUG_FUN__V70*/ meltfptr[69];;
		    /*epilog */

		    MELT_LOCATION ("warmelt-normal.melt:5610:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L25*/ meltfnum[19] = 0;
		    /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V70*/ meltfptr[69] = 0;
		  }
		  ;
		}
	      else
		{		/*^cond.else */

      /*_._IF___V69*/ meltfptr[67] = NULL;;
		}
	      ;

	      {
		MELT_LOCATION ("warmelt-normal.melt:5610:/ locexp");
		/*void */ (void) 0;
	      }
	      ;
	      /*^quasiblock */


	      /*epilog */

	      /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L24*/ meltfnum[18] = 0;
	      /*^clear */
	       /*clear *//*_._IF___V69*/ meltfptr[67] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*void */ (void) 0;
	      }
	      ;
	      /*epilog */
	    }

#endif /*MELT_HAVE_DEBUG */
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5611:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#IS_A__L26*/ meltfnum[19] =
		melt_is_instance_of ((melt_ptr_t)
				     ( /*_.NEXTB__V62*/ meltfptr[61]),
				     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[30])));;
	      MELT_LOCATION ("warmelt-normal.melt:5611:/ cond");
	      /*cond */ if ( /*_#IS_A__L26*/ meltfnum[19])	/*then */
		{
		  /*^cond.then */
		  /*_._IFELSE___V72*/ meltfptr[67] = ( /*nil */ NULL);;
		}
	      else
		{
		  MELT_LOCATION ("warmelt-normal.melt:5611:/ cond.else");

		  /*^block */
		  /*anyblock */
		  {




		    {
		      /*^locexp */
		      melt_assert_failed (("check nextb"),
					  ("warmelt-normal.melt")
					  ? ("warmelt-normal.melt") :
					  __FILE__,
					  (5611) ? (5611) : __LINE__,
					  __FUNCTION__);
		      ;
		    }
		    ;
		 /*clear *//*_._IFELSE___V72*/ meltfptr[67] = 0;
		    /*epilog */
		  }
		  ;
		}
	      ;
	      /*^compute */
	      /*_.IFCPP___V71*/ meltfptr[69] =
		/*_._IFELSE___V72*/ meltfptr[67];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5611:/ clear");
	       /*clear *//*_#IS_A__L26*/ meltfnum[19] = 0;
	      /*^clear */
	       /*clear *//*_._IFELSE___V72*/ meltfptr[67] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*_.IFCPP___V71*/ meltfptr[69] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5612:/ quasiblock");


	    /*^cond */
	    /*cond */ if (
			   /*ifisa */
			   melt_is_instance_of ((melt_ptr_t)
						( /*_.NEXTB__V62*/
						 meltfptr[61]),
						(melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[31])))
	      )			/*then */
	      {
		/*^cond.then */
		/*^getslot */
		{
		  melt_ptr_t slot = NULL, obj = NULL;
		  obj =
		    (melt_ptr_t) ( /*_.NEXTB__V62*/ meltfptr[61]) /*=obj*/ ;
		  melt_object_get_field (slot, obj, 0, "BINDER");
     /*_.NEXTSYM__V73*/ meltfptr[67] = slot;
		};
		;
	      }
	    else
	      {			/*^cond.else */

    /*_.NEXTSYM__V73*/ meltfptr[67] = NULL;;
	      }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
	    /*^compute */
   /*_#IS_A__L27*/ meltfnum[18] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.NEXTB__V62*/ meltfptr[61]),
				   (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[29])));;
	    MELT_LOCATION ("warmelt-normal.melt:5613:/ cond");
	    /*cond */ if ( /*_#IS_A__L27*/ meltfnum[18])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{


		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5614:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[1];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NLETREC__V55*/ meltfptr[54];
		    /*_.NRECLIST_FIND_LOCSYM__V75*/ meltfptr[74] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
				    tabval[32])),
				  (melt_ptr_t) ( /*_.NEXTSYM__V73*/
						meltfptr[67]),
				  (MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_.NEXTLOC__V74*/ meltfptr[73] =
		    /*_.NRECLIST_FIND_LOCSYM__V75*/ meltfptr[74];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5613:/ clear");
	       /*clear *//*_.NRECLIST_FIND_LOCSYM__V75*/ meltfptr[74] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

    /*_.NEXTLOC__V74*/ meltfptr[73] = NULL;;
	      }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5616:/ cond");
	    /*cond */ if (
			   /*ifisa */
			   melt_is_instance_of ((melt_ptr_t)
						( /*_.CURPAIRB__V61*/
						 meltfptr[60]),
						(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[29])))
	      )			/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @NPAIRB_HEAD",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.CURPAIRB__V61*/
						     meltfptr[60])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.CURPAIRB__V61*/ meltfptr[60]),
					(4),
					( /*_.CURNARG__V60*/ meltfptr[59]),
					"NPAIRB_HEAD");
		  ;
		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @NPAIRB_TAIL",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.CURPAIRB__V61*/
						     meltfptr[60])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.CURPAIRB__V61*/ meltfptr[60]),
					(5),
					( /*_.NEXTLOC__V74*/ meltfptr[73]),
					"NPAIRB_TAIL");
		  ;
		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @NCONSB_NLETREC",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.CURPAIRB__V61*/
						     meltfptr[60])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.CURPAIRB__V61*/ meltfptr[60]),
					(3),
					( /*_.NLETREC__V55*/ meltfptr[54]),
					"NCONSB_NLETREC");
		  ;
		  /*^touch */
		  meltgc_touch ( /*_.CURPAIRB__V61*/ meltfptr[60]);
		  ;
		  /*^touchobj */

		  melt_dbgtrace_written_object ( /*_.CURPAIRB__V61*/
						meltfptr[60], "put-fields");
		  ;
		  /*epilog */
		}
		;
	      }			/*noelse */
	    ;

#if MELT_HAVE_DEBUG
	    MELT_LOCATION ("warmelt-normal.melt:5619:/ cppif.then");
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		melt_dbgcounter++;
#endif
		;
	      }
	      ;

	      MELT_CHECK_SIGNAL ();
	      ;
     /*_#MELT_NEED_DBG__L28*/ meltfnum[19] =
		/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		;;
	      MELT_LOCATION ("warmelt-normal.melt:5619:/ cond");
	      /*cond */ if ( /*_#MELT_NEED_DBG__L28*/ meltfnum[19])	/*then */
		{
		  /*^cond.then */
		  /*^block */
		  /*anyblock */
		  {

       /*_#MELT_CALLCOUNT__L29*/ meltfnum[28] =
		      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		      meltcallcount	/* melt_callcount debugging */
#else
		      0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		      ;;

		    MELT_CHECK_SIGNAL ();
		    ;
		    MELT_LOCATION ("warmelt-normal.melt:5619:/ apply");
		    /*apply */
		    {
		      union meltparam_un argtab[5];
		      memset (&argtab, 0, sizeof (argtab));
		      /*^apply.arg */
		      argtab[0].meltbp_long =
			/*_#MELT_CALLCOUNT__L29*/ meltfnum[28];
		      /*^apply.arg */
		      argtab[1].meltbp_cstring = "warmelt-normal.melt";
		      /*^apply.arg */
		      argtab[2].meltbp_long = 5619;
		      /*^apply.arg */
		      argtab[3].meltbp_cstring =
			"normexp_list updated curpairb";
		      /*^apply.arg */
		      argtab[4].meltbp_aptr =
			(melt_ptr_t *) & /*_.CURPAIRB__V61*/ meltfptr[60];
		      /*_.MELT_DEBUG_FUN__V77*/ meltfptr[76] =
			melt_apply ((meltclosure_ptr_t)
				    (( /*!MELT_DEBUG_FUN */ meltfrout->
				      tabval[3])),
				    (melt_ptr_t) (( /*nil */ NULL)),
				    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				     MELTBPARSTR_PTR ""), argtab, "",
				    (union meltparam_un *) 0);
		    }
		    ;
		    /*_._IF___V76*/ meltfptr[74] =
		      /*_.MELT_DEBUG_FUN__V77*/ meltfptr[76];;
		    /*epilog */

		    MELT_LOCATION ("warmelt-normal.melt:5619:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L29*/ meltfnum[28] = 0;
		    /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V77*/ meltfptr[76] = 0;
		  }
		  ;
		}
	      else
		{		/*^cond.else */

      /*_._IF___V76*/ meltfptr[74] = NULL;;
		}
	      ;

	      {
		MELT_LOCATION ("warmelt-normal.melt:5619:/ locexp");
		/*void */ (void) 0;
	      }
	      ;
	      /*^quasiblock */


	      /*epilog */

	      /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L28*/ meltfnum[19] = 0;
	      /*^clear */
	       /*clear *//*_._IF___V76*/ meltfptr[74] = 0;
	    }

#else /*MELT_HAVE_DEBUG */
	    /*^cppif.else */
	    /*^block */
	    /*anyblock */
	    {


	      {
		/*^locexp */
		/*void */ (void) 0;
	      }
	      ;
	      /*epilog */
	    }

#endif /*MELT_HAVE_DEBUG */
	    ;

	    MELT_LOCATION ("warmelt-normal.melt:5612:/ clear");
	     /*clear *//*_.NEXTSYM__V73*/ meltfptr[67] = 0;
	    /*^clear */
	     /*clear *//*_#IS_A__L27*/ meltfnum[18] = 0;
	    /*^clear */
	     /*clear *//*_.NEXTLOC__V74*/ meltfptr[73] = 0;

	    MELT_LOCATION ("warmelt-normal.melt:5603:/ clear");
	     /*clear *//*_.CURPAIRB__V61*/ meltfptr[60] = 0;
	    /*^clear */
	     /*clear *//*_#plI__L18*/ meltfnum[14] = 0;
	    /*^clear */
	     /*clear *//*_.NEXTB__V62*/ meltfptr[61] = 0;
	    /*^clear */
	     /*clear *//*_.IFCPP___V65*/ meltfptr[63] = 0;
	    /*^clear */
	     /*clear *//*_.IFCPP___V67*/ meltfptr[62] = 0;
	    /*^clear */
	     /*clear *//*_.IFCPP___V71*/ meltfptr[69] = 0;
	    if ( /*_#NIX__L17*/ meltfnum[15] < 0)
	      break;
	  }			/* end  foreach_in_multiple meltcit3__EACHTUP */

	/*citerepilog */

	MELT_LOCATION ("warmelt-normal.melt:5600:/ clear");
	     /*clear *//*_.CURNARG__V60*/ meltfptr[59] = 0;
	/*^clear */
	     /*clear *//*_#NIX__L17*/ meltfnum[15] = 0;
      }				/*endciterblock FOREACH_IN_MULTIPLE */
      ;
      MELT_LOCATION ("warmelt-normal.melt:5624:/ quasiblock");


  /*_.FIRSTPAIRB__V78*/ meltfptr[76] =
	(melt_multiple_nth
	 ((melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]), (0)));;
      /*^compute */
  /*_.LASTPAIRB__V79*/ meltfptr[74] =
	(melt_multiple_nth
	 ((melt_ptr_t) ( /*_.CONSBINDTUP__V24*/ meltfptr[23]), (-2)));;

      MELT_CHECK_SIGNAL ();
      ;
      /*^compute */
  /*_#IS_A__L30*/ meltfnum[28] =
	melt_is_instance_of ((melt_ptr_t)
			     ( /*_.FIRSTPAIRB__V78*/ meltfptr[76]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[29])));;
      MELT_LOCATION ("warmelt-normal.melt:5627:/ cond");
      /*cond */ if ( /*_#IS_A__L30*/ meltfnum[28])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    MELT_LOCATION ("warmelt-normal.melt:5629:/ getslot");
	    {
	      melt_ptr_t slot = NULL, obj = NULL;
	      obj =
		(melt_ptr_t) ( /*_.FIRSTPAIRB__V78*/ meltfptr[76]) /*=obj*/ ;
	      melt_object_get_field (slot, obj, 0, "BINDER");
     /*_.BINDER__V81*/ meltfptr[73] = slot;
	    };
	    ;
	    /*_.FIRSTPAIRSYMB__V80*/ meltfptr[67] =
	      /*_.BINDER__V81*/ meltfptr[73];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5627:/ clear");
	      /*clear *//*_.BINDER__V81*/ meltfptr[73] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

   /*_.FIRSTPAIRSYMB__V80*/ meltfptr[67] = NULL;;
	}
      ;

      MELT_CHECK_SIGNAL ();
      ;
      /*^compute */
  /*_#IS_A__L31*/ meltfnum[19] =
	melt_is_instance_of ((melt_ptr_t)
			     ( /*_.LASTPAIRB__V79*/ meltfptr[74]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[29])));;
      MELT_LOCATION ("warmelt-normal.melt:5631:/ cond");
      /*cond */ if ( /*_#IS_A__L31*/ meltfnum[19])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    MELT_LOCATION ("warmelt-normal.melt:5633:/ getslot");
	    {
	      melt_ptr_t slot = NULL, obj = NULL;
	      obj =
		(melt_ptr_t) ( /*_.LASTPAIRB__V79*/ meltfptr[74]) /*=obj*/ ;
	      melt_object_get_field (slot, obj, 0, "BINDER");
     /*_.BINDER__V83*/ meltfptr[61] = slot;
	    };
	    ;
	    /*_.LASTPAIRSYMB__V82*/ meltfptr[60] =
	      /*_.BINDER__V83*/ meltfptr[61];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5631:/ clear");
	      /*clear *//*_.BINDER__V83*/ meltfptr[61] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

   /*_.LASTPAIRSYMB__V82*/ meltfptr[60] = NULL;;
	}
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5635:/ cond");
      /*cond */ if ( /*_.FIRSTPAIRSYMB__V80*/ meltfptr[67])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5636:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[1];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V55*/ meltfptr[54];
	      /*_.NRECLIST_FIND_LOCSYM__V85*/ meltfptr[62] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
			      tabval[32])),
			    (melt_ptr_t) ( /*_.FIRSTPAIRSYMB__V80*/
					  meltfptr[67]), (MELTBPARSTR_PTR ""),
			    argtab, "", (union meltparam_un *) 0);
	    }
	    ;
	    /*_.FIRSTPAIRLOC__V84*/ meltfptr[63] =
	      /*_.NRECLIST_FIND_LOCSYM__V85*/ meltfptr[62];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5635:/ clear");
	      /*clear *//*_.NRECLIST_FIND_LOCSYM__V85*/ meltfptr[62] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

   /*_.FIRSTPAIRLOC__V84*/ meltfptr[63] = NULL;;
	}
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5638:/ cond");
      /*cond */ if ( /*_.LASTPAIRSYMB__V82*/ meltfptr[60])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5639:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[1];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V55*/ meltfptr[54];
	      /*_.NRECLIST_FIND_LOCSYM__V87*/ meltfptr[73] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
			      tabval[32])),
			    (melt_ptr_t) ( /*_.LASTPAIRSYMB__V82*/
					  meltfptr[60]), (MELTBPARSTR_PTR ""),
			    argtab, "", (union meltparam_un *) 0);
	    }
	    ;
	    /*_.LASTPAIRLOC__V86*/ meltfptr[69] =
	      /*_.NRECLIST_FIND_LOCSYM__V87*/ meltfptr[73];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5638:/ clear");
	      /*clear *//*_.NRECLIST_FIND_LOCSYM__V87*/ meltfptr[73] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

   /*_.LASTPAIRLOC__V86*/ meltfptr[69] = NULL;;
	}
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5641:/ cond");
      /*cond */ if (
		     /*ifisa */
		     melt_is_instance_of ((melt_ptr_t)
					  ( /*_.CONSLISTBIND__V26*/
					   meltfptr[25]),
					  (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LIST_BINDING */ meltfrout->tabval[20])))
	)			/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    /*^putslot */
	    /*putslot */
	    melt_assertmsg ("checkobj putslot  _ @NLISTB_FIRST",
			    melt_magic_discr ((melt_ptr_t)
					      ( /*_.CONSLISTBIND__V26*/
					       meltfptr[25])) ==
			    MELTOBMAG_OBJECT);
	    melt_putfield_object (( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
				  (4),
				  ( /*_.FIRSTPAIRLOC__V84*/ meltfptr[63]),
				  "NLISTB_FIRST");
	    ;
	    /*^putslot */
	    /*putslot */
	    melt_assertmsg ("checkobj putslot  _ @NLISTB_LAST",
			    melt_magic_discr ((melt_ptr_t)
					      ( /*_.CONSLISTBIND__V26*/
					       meltfptr[25])) ==
			    MELTOBMAG_OBJECT);
	    melt_putfield_object (( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
				  (5), ( /*_.LASTPAIRLOC__V86*/ meltfptr[69]),
				  "NLISTB_LAST");
	    ;
	    /*^putslot */
	    /*putslot */
	    melt_assertmsg ("checkobj putslot  _ @NCONSB_NLETREC",
			    melt_magic_discr ((melt_ptr_t)
					      ( /*_.CONSLISTBIND__V26*/
					       meltfptr[25])) ==
			    MELTOBMAG_OBJECT);
	    melt_putfield_object (( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
				  (3), ( /*_.NLETREC__V55*/ meltfptr[54]),
				  "NCONSB_NLETREC");
	    ;
	    /*^touch */
	    meltgc_touch ( /*_.CONSLISTBIND__V26*/ meltfptr[25]);
	    ;
	    /*^touchobj */

	    melt_dbgtrace_written_object ( /*_.CONSLISTBIND__V26*/
					  meltfptr[25], "put-fields");
	    ;
	    /*epilog */
	  }
	  ;
	}			/*noelse */
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5645:/ locexp");
	melt_put_int ((melt_ptr_t) ( /*_.CONSLISTBIND__V26*/ meltfptr[25]),
		      ( /*_#NBARGS__L6*/ meltfnum[4]));
      }
      ;

#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5646:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L32*/ meltfnum[18] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5646:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L32*/ meltfnum[18])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L33*/ meltfnum[14] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5646:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[5];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L33*/ meltfnum[14];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5646;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_list updated conslistbind";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.CONSLISTBIND__V26*/ meltfptr[25];
		/*_.MELT_DEBUG_FUN__V89*/ meltfptr[62] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V88*/ meltfptr[61] =
		/*_.MELT_DEBUG_FUN__V89*/ meltfptr[62];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5646:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L33*/ meltfnum[14] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V89*/ meltfptr[62] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V88*/ meltfptr[61] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5646:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L32*/ meltfnum[18] = 0;
	/*^clear */
	      /*clear *//*_._IF___V88*/ meltfptr[61] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_LOCATION ("warmelt-normal.melt:5624:/ clear");
	    /*clear *//*_.FIRSTPAIRB__V78*/ meltfptr[76] = 0;
      /*^clear */
	    /*clear *//*_.LASTPAIRB__V79*/ meltfptr[74] = 0;
      /*^clear */
	    /*clear *//*_#IS_A__L30*/ meltfnum[28] = 0;
      /*^clear */
	    /*clear *//*_.FIRSTPAIRSYMB__V80*/ meltfptr[67] = 0;
      /*^clear */
	    /*clear *//*_#IS_A__L31*/ meltfnum[19] = 0;
      /*^clear */
	    /*clear *//*_.LASTPAIRSYMB__V82*/ meltfptr[60] = 0;
      /*^clear */
	    /*clear *//*_.FIRSTPAIRLOC__V84*/ meltfptr[63] = 0;
      /*^clear */
	    /*clear *//*_.LASTPAIRLOC__V86*/ meltfptr[69] = 0;

      MELT_CHECK_SIGNAL ();
      ;
  /*_#NULL__L34*/ meltfnum[14] =
	(( /*_.NBINDINGS__V36*/ meltfptr[35]) == NULL);;
      MELT_LOCATION ("warmelt-normal.melt:5649:/ cond");
      /*cond */ if ( /*_#NULL__L34*/ meltfnum[14])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    MELT_LOCATION ("warmelt-normal.melt:5650:/ compute");
	    /*_.NBINDINGS__V36*/ meltfptr[35] =
	      /*_.SETQ___V91*/ meltfptr[62] =
	      /*_.CLISTBINDLIST__V59*/ meltfptr[58];;
	    /*_._IFELSE___V90*/ meltfptr[73] = /*_.SETQ___V91*/ meltfptr[62];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5649:/ clear");
	      /*clear *//*_.SETQ___V91*/ meltfptr[62] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

	  /*^block */
	  /*anyblock */
	  {




	    {
	      MELT_LOCATION ("warmelt-normal.melt:5651:/ locexp");
	      meltgc_append_list ((melt_ptr_t)
				  ( /*_.NBINDINGS__V36*/ meltfptr[35]),
				  (melt_ptr_t) ( /*_.CLISTBIND__V56*/
						meltfptr[55]));
	    }
	    ;
	      /*clear *//*_._IFELSE___V90*/ meltfptr[73] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;

#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5653:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L35*/ meltfnum[18] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5653:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L35*/ meltfnum[18])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L36*/ meltfnum[28] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5653:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[9];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L36*/ meltfnum[28];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5653;
		/*^apply.arg */
		argtab[3].meltbp_cstring = "normexp_list final nletrec=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NLETREC__V55*/ meltfptr[54];
		/*^apply.arg */
		argtab[5].meltbp_cstring = " clistlocc=";
		/*^apply.arg */
		argtab[6].meltbp_aptr =
		  (melt_ptr_t *) & /*_.CLISTLOCC__V57*/ meltfptr[56];
		/*^apply.arg */
		argtab[7].meltbp_cstring = " nbindings=";
		/*^apply.arg */
		argtab[8].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NBINDINGS__V36*/ meltfptr[35];
		/*_.MELT_DEBUG_FUN__V93*/ meltfptr[76] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V92*/ meltfptr[61] =
		/*_.MELT_DEBUG_FUN__V93*/ meltfptr[76];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5653:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L36*/ meltfnum[28] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V93*/ meltfptr[76] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V92*/ meltfptr[61] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5653:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L35*/ meltfnum[18] = 0;
	/*^clear */
	      /*clear *//*_._IF___V92*/ meltfptr[61] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5655:/ quasiblock");


      /*_._RETVAL___V1*/ meltfptr[0] = /*_.CLISTLOCC__V57*/ meltfptr[56];;
      MELT_LOCATION ("warmelt-normal.melt:5655:/ putxtraresult");
      if (!meltxrestab_ || !meltxresdescr_)
	goto meltlabend_rout;
      if (meltxresdescr_[0] != MELTBPAR_PTR)
	goto meltlabend_rout;
      if (meltxrestab_[0].meltbp_aptr)
	*(meltxrestab_[0].meltbp_aptr) =
	  (melt_ptr_t) ( /*_.NBINDINGS__V36*/ meltfptr[35]);
      ;
      /*^finalreturn */
      ;
      /*finalret */ goto meltlabend_rout;
      /*_.LETREC___V52*/ meltfptr[48] = /*_.RETURN___V94*/ meltfptr[74];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5573:/ clear");
	    /*clear *//*_.CRECLOCC__V53*/ meltfptr[47] = 0;
      /*^clear */
	    /*clear *//*_.NBDY__V54*/ meltfptr[53] = 0;
      /*^clear */
	    /*clear *//*_.NLETREC__V55*/ meltfptr[54] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBIND__V56*/ meltfptr[55] = 0;
      /*^clear */
	    /*clear *//*_.CLISTLOCC__V57*/ meltfptr[56] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBIND__V58*/ meltfptr[57] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBINDLIST__V59*/ meltfptr[58] = 0;
      /*^clear */
	    /*clear *//*_.CRECLOCC__V53*/ meltfptr[47] = 0;
      /*^clear */
	    /*clear *//*_.NBDY__V54*/ meltfptr[53] = 0;
      /*^clear */
	    /*clear *//*_.NLETREC__V55*/ meltfptr[54] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBIND__V56*/ meltfptr[55] = 0;
      /*^clear */
	    /*clear *//*_.CLISTLOCC__V57*/ meltfptr[56] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBIND__V58*/ meltfptr[57] = 0;
      /*^clear */
	    /*clear *//*_.CLISTBINDLIST__V59*/ meltfptr[58] = 0;
      /*^clear */
	    /*clear *//*_#NULL__L34*/ meltfnum[14] = 0;
      /*^clear */
	    /*clear *//*_._IFELSE___V90*/ meltfptr[73] = 0;
      /*^clear */
	    /*clear *//*_.RETURN___V94*/ meltfptr[74] = 0;
    }				/*end multiallocblock */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    /*_.MULTI___V34*/ meltfptr[29] = /*_.LETREC___V52*/ meltfptr[48];;

    MELT_LOCATION ("warmelt-normal.melt:5554:/ clear");
	   /*clear *//*_.IFCPP___V39*/ meltfptr[37] = 0;
    /*^clear */
	   /*clear *//*_.LIST_APPEND2LIST__V41*/ meltfptr[36] = 0;
    /*^clear */
	   /*clear *//*_.LETREC___V52*/ meltfptr[48] = 0;

    /*^clear */
	   /*clear *//*_.NBINDINGS__V36*/ meltfptr[35] = 0;
    /*_.LET___V14*/ meltfptr[12] = /*_.MULTI___V34*/ meltfptr[29];;

    MELT_LOCATION ("warmelt-normal.melt:5526:/ clear");
	   /*clear *//*_.LOCA_LOCATION__V15*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.SLOC__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.SARGS__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.CSYMREC__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.CSYMLIST__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.NORMDISCRPAIR__V20*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_.NORMDISCRLIST__V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.NEWENV__V22*/ meltfptr[21] = 0;
    /*^clear */
	   /*clear *//*_#NBARGS__L6*/ meltfnum[4] = 0;
    /*^clear */
	   /*clear *//*_.PREPARE_CONSTRUCTOR_BINDING__V23*/ meltfptr[22] = 0;
    /*^clear */
	   /*clear *//*_.CONSBINDTUP__V24*/ meltfptr[23] = 0;
    /*^clear */
	   /*clear *//*_.RESBINDINGS__V25*/ meltfptr[24] = 0;
    /*^clear */
	   /*clear *//*_.CONSLISTBIND__V26*/ meltfptr[25] = 0;
    /*^clear */
	   /*clear *//*_.LAMBDA___V27*/ meltfptr[26] = 0;
    /*^clear */
	   /*clear *//*_.TUPRECLOC__V29*/ meltfptr[28] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V32*/ meltfptr[30] = 0;
    /*^clear */
	   /*clear *//*_.MULTI___V34*/ meltfptr[29] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5521:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V14*/ meltfptr[12];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5521:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V6*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.LET___V14*/ meltfptr[12] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMEXP_LIST", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_99_WARMELTmiNORMAL_NORMEXP_LIST */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un * meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32_melt =
    melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 11
    melt_ptr_t mcfr_varptr[11];
#define MELTFRAM_NBVARNUM 4
    long mcfr_varnum[4];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32_st *)
	meltfirstargp_;
      /* use arguments meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 11; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 nbval 11*/
  meltfram__.mcfr_nbvar = 11 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("LAMBDA_cl32", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5539:/ getarg");
 /*_.CURBIND__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_LONG)
    goto meltlab_endgetargs;
 /*_#BINDIX__L1*/ meltfnum[0] = meltxargtab_[0].meltbp_long;
  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5540:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L2*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5540:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L2*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L3*/ meltfnum[2] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5540:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L3*/ meltfnum[2];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5540;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_list curbind";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CURBIND__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V4*/ meltfptr[3] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V3*/ meltfptr[2] =
	      /*_.MELT_DEBUG_FUN__V4*/ meltfptr[3];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5540:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L3*/ meltfnum[2] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V4*/ meltfptr[3] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V3*/ meltfptr[2] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5540:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L2*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V3*/ meltfptr[2] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5541:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[2] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CURBIND__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5541:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[2])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V6*/ meltfptr[2] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5541:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check curbind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5541) ? (5541) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V6*/ meltfptr[2] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V5*/ meltfptr[3] = /*_._IFELSE___V6*/ meltfptr[2];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5541:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[2] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V6*/ meltfptr[2] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V5*/ meltfptr[3] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5543:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CURBIND__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CURBIND__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "NCONSB_LOC");
   /*_.NCONSB_LOC__V7*/ meltfptr[2] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NCONSB_LOC__V7*/ meltfptr[2] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5543:/ cond");
    /*cond */ if ( /*_.NCONSB_LOC__V7*/ meltfptr[2])	/*then */
      {
	/*^cond.then */
	/*_.OR___V8*/ meltfptr[7] = /*_.NCONSB_LOC__V7*/ meltfptr[2];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5543:/ cond.else");

	/*_.OR___V8*/ meltfptr[7] = ( /*~SLOC */ meltfclos->tabval[0]);;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5545:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CURBIND__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[4])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CURBIND__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 0, "BINDER");
   /*_.BINDER__V9*/ meltfptr[8] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.BINDER__V9*/ meltfptr[8] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5542:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NREP_LOCSYMOCC */
					     meltfrout->tabval[2])), (4),
			      "CLASS_NREP_LOCSYMOCC");
  /*_.INST__V11*/ meltfptr[10] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V11*/ meltfptr[10])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V11*/ meltfptr[10]), (0),
			  ( /*_.OR___V8*/ meltfptr[7]), "NREP_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NOCC_CTYP",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V11*/ meltfptr[10])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V11*/ meltfptr[10]), (2),
			  (( /*!CTYPE_VALUE */ meltfrout->tabval[3])),
			  "NOCC_CTYP");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NOCC_SYMB",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V11*/ meltfptr[10])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V11*/ meltfptr[10]), (1),
			  ( /*_.BINDER__V9*/ meltfptr[8]), "NOCC_SYMB");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NOCC_BIND",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V11*/ meltfptr[10])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V11*/ meltfptr[10]), (3),
			  ( /*_.CURBIND__V2*/ meltfptr[1]), "NOCC_BIND");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V11*/ meltfptr[10],
				  "newly made instance");
    ;
    /*_.INST___V10*/ meltfptr[9] = /*_.INST__V11*/ meltfptr[10];;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5539:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.INST___V10*/ meltfptr[9];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5539:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V5*/ meltfptr[3] = 0;
    /*^clear */
	   /*clear *//*_.NCONSB_LOC__V7*/ meltfptr[2] = 0;
    /*^clear */
	   /*clear *//*_.OR___V8*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_.BINDER__V9*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.INST___V10*/ meltfptr[9] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("LAMBDA_cl32", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_100_WARMELTmiNORMAL_LAMBDA_cl32 */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION
  (meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_,
   const melt_argdescr_cell_t meltxargdescr_[],
   union meltparam_un * meltxargtab_,
   const melt_argdescr_cell_t meltxresdescr_[],
   union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 87
    melt_ptr_t mcfr_varptr[87];
#define MELTFRAM_NBVARNUM 30
    long mcfr_varnum[30];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct
	 meltframe_meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION_st
	 *) meltfirstargp_;
      /* use arguments meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 87; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION nbval 87*/
  meltfram__.mcfr_nbvar = 87 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMEXP_ARITHMETIC_VARIADIC_OPERATION", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5661:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.PSLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.PSLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5662:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5662:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5662:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5662;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normexp_arithmetic_variadic_operation recv";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5662:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5662:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5663:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_ARITHMETIC_VARIADIC_OPERATION */ meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5663:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5663:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5663) ? (5663) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5663:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5664:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:5664:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[10] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5664:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check env"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5664) ? (5664) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[5] = /*_._IFELSE___V11*/ meltfptr[10];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5664:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5665:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L5*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[3])));;
      MELT_LOCATION ("warmelt-normal.melt:5665:/ cond");
      /*cond */ if ( /*_#IS_A__L5*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V13*/ meltfptr[12] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5665:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5665) ? (5665) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V12*/ meltfptr[10] = /*_._IFELSE___V13*/ meltfptr[12];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5665:/ clear");
	     /*clear *//*_#IS_A__L5*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V12*/ meltfptr[10] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5666:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.LOCA_LOCATION__V14*/ meltfptr[12] = slot;
    };
    ;
    /*^cond */
    /*cond */ if ( /*_.LOCA_LOCATION__V14*/ meltfptr[12])	/*then */
      {
	/*^cond.then */
	/*_.SLOC__V15*/ meltfptr[14] = /*_.LOCA_LOCATION__V14*/ meltfptr[12];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5666:/ cond.else");

	/*_.SLOC__V15*/ meltfptr[14] = /*_.PSLOC__V5*/ meltfptr[4];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5667:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
  /*_.SARGS__V16*/ meltfptr[15] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:5668:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 3, "SARITHVAR_NEUTRAL");
  /*_.SNEUTRAL__V17*/ meltfptr[16] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:5669:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 4, "SARITHVAR_PRIMITIVE");
  /*_.SPRIMITIVE__V18*/ meltfptr[17] = slot;
    };
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5671:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L6*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5671:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L6*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L7*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5671:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L7*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5671;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normexp_arithmetic_variadic_operation sprimitive=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SPRIMITIVE__V18*/ meltfptr[17];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " sneutral=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SNEUTRAL__V17*/ meltfptr[16];
	      /*_.MELT_DEBUG_FUN__V20*/ meltfptr[19] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V19*/ meltfptr[18] =
	      /*_.MELT_DEBUG_FUN__V20*/ meltfptr[19];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5671:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L7*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V20*/ meltfptr[19] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V19*/ meltfptr[18] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5671:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L6*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V19*/ meltfptr[18] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5672:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L8*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t)
			     ( /*_.SPRIMITIVE__V18*/ meltfptr[17]),
			     (melt_ptr_t) (( /*!CLASS_PRIMITIVE */ meltfrout->
					    tabval[4])));;
      MELT_LOCATION ("warmelt-normal.melt:5672:/ cond");
      /*cond */ if ( /*_#IS_A__L8*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V22*/ meltfptr[18] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5672:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check sprimitive"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5672) ? (5672) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V22*/ meltfptr[18] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V21*/ meltfptr[19] = /*_._IFELSE___V22*/ meltfptr[18];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5672:/ clear");
	     /*clear *//*_#IS_A__L8*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V22*/ meltfptr[18] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V21*/ meltfptr[19] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5673:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_INTEGERBOX__L9*/ meltfnum[0] =
	(melt_magic_discr ((melt_ptr_t) ( /*_.SNEUTRAL__V17*/ meltfptr[16]))
	 == MELTOBMAG_INT);;
      MELT_LOCATION ("warmelt-normal.melt:5673:/ cond");
      /*cond */ if ( /*_#IS_INTEGERBOX__L9*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V24*/ meltfptr[23] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5673:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check sneutral"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5673) ? (5673) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V24*/ meltfptr[23] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V23*/ meltfptr[18] = /*_._IFELSE___V24*/ meltfptr[23];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5673:/ clear");
	     /*clear *//*_#IS_INTEGERBOX__L9*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V24*/ meltfptr[23] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V23*/ meltfptr[18] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5676:/ quasiblock");


    /*^multiapply */
    /*multiapply 4args, 1x.res */
    {
      union meltparam_un argtab[3];

      union meltparam_un restab[1];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[14];
      /*^multiapply.xres */
      restab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.NBINDINGS__V26*/ meltfptr[25];
      /*^multiapply.appl */
      /*_.NARGS__V25*/ meltfptr[23] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_TUPLE */ meltfrout->tabval[5])),
		    (melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		    argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5679:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L10*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5679:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L11*/ meltfnum[0] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5679:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L11*/ meltfnum[0];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5679;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normexp_arithmetic_variadic_operation nargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NARGS__V25*/ meltfptr[23];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nbindings=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBINDINGS__V26*/ meltfptr[25];
	      /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V27*/ meltfptr[26] =
	      /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5679:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[0] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V27*/ meltfptr[26] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5679:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V27*/ meltfptr[26] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5680:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L12*/ meltfnum[0] =
	(( /*_.NBINDINGS__V26*/ meltfptr[25]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBINDINGS__V26*/ meltfptr[25])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:5680:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L12*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V30*/ meltfptr[26] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5680:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nbindings"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5680) ? (5680) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V30*/ meltfptr[26] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V29*/ meltfptr[27] = /*_._IFELSE___V30*/ meltfptr[26];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5680:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L12*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V30*/ meltfptr[26] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V29*/ meltfptr[27] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NARGS__V25*/ meltfptr[23]);
      for ( /*_#ARGIX__L13*/ meltfnum[1] = 0;
	   ( /*_#ARGIX__L13*/ meltfnum[1] >= 0)
	   && ( /*_#ARGIX__L13*/ meltfnum[1] < meltcit1__EACHTUP_ln);
	/*_#ARGIX__L13*/ meltfnum[1]++)
	{
	  /*_.CURNARG__V31*/ meltfptr[26] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.NARGS__V25*/ meltfptr[23]),
			       /*_#ARGIX__L13*/ meltfnum[1]);



	  MELT_LOCATION ("warmelt-normal.melt:5685:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^msend */
	  /*msend */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];
	    /*_.ARGCTYP__V33*/ meltfptr[32] =
	      meltgc_send ((melt_ptr_t) ( /*_.CURNARG__V31*/ meltfptr[26]),
			   (melt_ptr_t) (( /*!GET_CTYPE */ meltfrout->
					  tabval[6])), (MELTBPARSTR_PTR ""),
			   argtab, "", (union meltparam_un *) 0);
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
  /*_#exeq__L14*/ meltfnum[0] =
	    (( /*_.ARGCTYP__V33*/ meltfptr[32]) !=
	     (( /*!CTYPE_LONG */ meltfrout->tabval[7])));;
	  MELT_LOCATION ("warmelt-normal.melt:5687:/ cond");
	  /*cond */ if ( /*_#exeq__L14*/ meltfnum[0])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {


		{
		  MELT_LOCATION ("warmelt-normal.melt:5688:/ locexp");
		  /* error_plain */
		    melt_error_str ((melt_ptr_t)
				    ( /*_.SLOC__V15*/ meltfptr[14]),
				    ("variadic arithmetic operation requires only :long arguments"),
				    (melt_ptr_t) 0);
		}
		;
		MELT_LOCATION ("warmelt-normal.melt:5689:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = NULL;;

		{
		  MELT_LOCATION ("warmelt-normal.melt:5689:/ locexp");
		  /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
		  if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
		    melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
		  ;
		}
		;
		/*^finalreturn */
		;
		/*finalret */ goto meltlabend_rout;
		MELT_LOCATION ("warmelt-normal.melt:5687:/ quasiblock");


		/*_.PROGN___V36*/ meltfptr[35] =
		  /*_.RETURN___V35*/ meltfptr[34];;
		/*^compute */
		/*_._IF___V34*/ meltfptr[33] =
		  /*_.PROGN___V36*/ meltfptr[35];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5687:/ clear");
	      /*clear *//*_.RETURN___V35*/ meltfptr[34] = 0;
		/*^clear */
	      /*clear *//*_.PROGN___V36*/ meltfptr[35] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_._IF___V34*/ meltfptr[33] = NULL;;
	    }
	  ;
	  /*^compute */
	  /*_.LET___V32*/ meltfptr[31] = /*_._IF___V34*/ meltfptr[33];;

	  MELT_LOCATION ("warmelt-normal.melt:5685:/ clear");
	    /*clear *//*_.ARGCTYP__V33*/ meltfptr[32] = 0;
	  /*^clear */
	    /*clear *//*_#exeq__L14*/ meltfnum[0] = 0;
	  /*^clear */
	    /*clear *//*_._IF___V34*/ meltfptr[33] = 0;
	  if ( /*_#ARGIX__L13*/ meltfnum[1] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5682:/ clear");
	    /*clear *//*_.CURNARG__V31*/ meltfptr[26] = 0;
      /*^clear */
	    /*clear *//*_#ARGIX__L13*/ meltfnum[1] = 0;
      /*^clear */
	    /*clear *//*_.LET___V32*/ meltfptr[31] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5691:/ block");
    /*anyblock */
    {

      /*^objgoto */
      /*objgoto */ goto mtch1_0;
      ;

    /*objlabel */ mtch1_0:;
      MELT_LOCATION ("warmelt-normal.melt:5693:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_8;
      meltlab_count_8++;
      debugeprintf
	("objlabel_8 mtch1_*CLASS_NORMTESTER_TUPLE/f032502. *CLASS_OBJLABELINSTR/3023d2ca. %ld",
	 meltlab_count_8);
#endif

      ;
      /*^clear */
	    /*clear *//*_.NTHCOMP__V37*/ meltfptr[34] = 0;
      /*^cond */
      /*cond */ if (
		     /*normtestertuple */
	(melt_magic_discr ((melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15])) ==
	 MELTOBMAG_MULTIPLE
	 && ((meltmultiple_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]))->nbval == 1))	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    /*^compute */
    /*_.NTHCOMP__V37*/ meltfptr[34] =
	      /*unsafenthcomp */
	      ((meltmultiple_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]))->
	      tabval[0];;
	    /*^objgoto */
	    /*objgoto */ goto mtch1_1;
	    ;
	  }
	  ;
	}
      else
	{			/*^cond.else */

	  /*^block */
	  /*anyblock */
	  {

	    MELT_LOCATION ("warmelt-normal.melt:5708:/ objgoto");
	    /*objgoto */ goto mtch1_2;
	    ;
	  }
	  ;
	}
      ;

    /*objlabel */ mtch1_1:;
      MELT_LOCATION ("warmelt-normal.melt:5693:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_9;
      meltlab_count_9++;
      debugeprintf
	("objlabel_9 mtch1_*CLASS_NORMTESTER_SUCCESS/3addf0ad. *CLASS_OBJLABELINSTR/f9a5d3e. %ld",
	 meltlab_count_9);
#endif

      ;
      /*^quasiblock */


      /*_.S0__V38*/ meltfptr[35] = /*_.NTHCOMP__V37*/ meltfptr[34];;
      MELT_LOCATION ("warmelt-normal.melt:5694:/ quasiblock");



      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5697:/ blockmultialloc");
      /*multiallocblock */
      {
	struct meltletrec_1_st
	{
	  struct MELT_MULTIPLE_STRUCT (2) rtup_0__TUPLREC__x12;
	  long meltletrec_1_endgap;
	} *meltletrec_1_ptr = 0;
	meltletrec_1_ptr =
	  (struct meltletrec_1_st *)
	  meltgc_allocate (sizeof (struct meltletrec_1_st), 0);
	/*^blockmultialloc.initfill */
	/*inimult rtup_0__TUPLREC__x12 */
 /*_.TUPLREC___V40*/ meltfptr[33] =
	  (melt_ptr_t) & meltletrec_1_ptr->rtup_0__TUPLREC__x12;
	meltletrec_1_ptr->rtup_0__TUPLREC__x12.discr =
	  (meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_MULTIPLE))));
	meltletrec_1_ptr->rtup_0__TUPLREC__x12.nbval = 2;


	/*^putuple */
	/*putupl#41 */
	melt_assertmsg ("putupl [:5697] #41 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V40*/ meltfptr[33]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5697] #41 checkoff",
			(0 >= 0
			 && 0 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V40*/
						meltfptr[33]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V40*/ meltfptr[33]))->
	  tabval[0] = (melt_ptr_t) ( /*_.SNEUTRAL__V17*/ meltfptr[16]);
	;
	/*^putuple */
	/*putupl#42 */
	melt_assertmsg ("putupl [:5697] #42 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V40*/ meltfptr[33]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5697] #42 checkoff",
			(1 >= 0
			 && 1 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V40*/
						meltfptr[33]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V40*/ meltfptr[33]))->
	  tabval[1] = (melt_ptr_t) ( /*_.S0__V38*/ meltfptr[35]);
	;
	/*^touch */
	meltgc_touch ( /*_.TUPLREC___V40*/ meltfptr[33]);
	;
	/*_.TUPLE___V39*/ meltfptr[32] = /*_.TUPLREC___V40*/ meltfptr[33];;
	/*epilog */

	MELT_LOCATION ("warmelt-normal.melt:5697:/ clear");
	     /*clear *//*_.TUPLREC___V40*/ meltfptr[33] = 0;
	/*^clear */
	     /*clear *//*_.TUPLREC___V40*/ meltfptr[33] = 0;
      }				/*end multiallocblock */
      ;
      MELT_LOCATION ("warmelt-normal.melt:5694:/ quasiblock");


      /*^rawallocobj */
      /*rawallocobj */
      {
	melt_ptr_t newobj = 0;
	melt_raw_object_create (newobj,
				(melt_ptr_t) (( /*!CLASS_SOURCE_PRIMITIVE */
					       meltfrout->tabval[8])), (4),
				"CLASS_SOURCE_PRIMITIVE");
   /*_.INST__V42*/ meltfptr[41] =
	  newobj;
      };
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LOCA_LOCATION",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V42*/ meltfptr[41])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V42*/ meltfptr[41]), (1),
			    ( /*_.SLOC__V15*/ meltfptr[14]), "LOCA_LOCATION");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SPRIM_OPER",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V42*/ meltfptr[41])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V42*/ meltfptr[41]), (3),
			    ( /*_.SPRIMITIVE__V18*/ meltfptr[17]),
			    "SPRIM_OPER");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SARGOP_ARGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V42*/ meltfptr[41])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V42*/ meltfptr[41]), (2),
			    ( /*_.TUPLE___V39*/ meltfptr[32]), "SARGOP_ARGS");
      ;
      /*^touchobj */

      melt_dbgtrace_written_object ( /*_.INST__V42*/ meltfptr[41],
				    "newly made instance");
      ;
      /*_.SPRIM__V41*/ meltfptr[33] = /*_.INST__V42*/ meltfptr[41];;

#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5699:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L15*/ meltfnum[0] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5699:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L15*/ meltfnum[0])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L16*/ meltfnum[15] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5699:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[5];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L16*/ meltfnum[15];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5699;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation unary sprim=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.SPRIM__V41*/ meltfptr[33];
		/*_.MELT_DEBUG_FUN__V44*/ meltfptr[43] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V43*/ meltfptr[42] =
		/*_.MELT_DEBUG_FUN__V44*/ meltfptr[43];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5699:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L16*/ meltfnum[15] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V44*/ meltfptr[43] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V43*/ meltfptr[42] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5699:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L15*/ meltfnum[0] = 0;
	/*^clear */
	      /*clear *//*_._IF___V43*/ meltfptr[42] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5700:/ quasiblock");


      /*^multiapply */
      /*multiapply 4args, 1x.res */
      {
	union meltparam_un argtab[3];

	union meltparam_un restab[1];
	memset (&restab, 0, sizeof (restab));
	memset (&argtab, 0, sizeof (argtab));
	/*^multiapply.arg */
	argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];	/*^multiapply.arg */
	argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
	argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.PSLOC__V5*/ meltfptr[4];
	/*^multiapply.xres */
	restab[0].meltbp_aptr =
	  (melt_ptr_t *) & /*_.NREALBINDINGS__V46*/ meltfptr[42];
	/*^multiapply.appl */
	/*_.NREALARGS__V45*/ meltfptr[43] =
	  melt_apply ((meltclosure_ptr_t)
		      (( /*!NORMEXP_PRIMITIVE */ meltfrout->tabval[9])),
		      (melt_ptr_t) ( /*_.SPRIM__V41*/ meltfptr[33]),
		      (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		      argtab, (MELTBPARSTR_PTR ""), restab);
      }
      ;
      /*^quasiblock */



#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5703:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L17*/ meltfnum[15] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5703:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L17*/ meltfnum[15])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L18*/ meltfnum[0] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5703:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[7];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L18*/ meltfnum[0];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5703;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation unary result nrealargs=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALARGS__V45*/ meltfptr[43];
		/*^apply.arg */
		argtab[5].meltbp_cstring = " nrealbindings=";
		/*^apply.arg */
		argtab[6].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALBINDINGS__V46*/ meltfptr[42];
		/*_.MELT_DEBUG_FUN__V48*/ meltfptr[47] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V47*/ meltfptr[46] =
		/*_.MELT_DEBUG_FUN__V48*/ meltfptr[47];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5703:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L18*/ meltfnum[0] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V48*/ meltfptr[47] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V47*/ meltfptr[46] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5703:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L17*/ meltfnum[15] = 0;
	/*^clear */
	      /*clear *//*_._IF___V47*/ meltfptr[46] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5705:/ quasiblock");


      /*_._RETVAL___V1*/ meltfptr[0] = /*_.NREALARGS__V45*/ meltfptr[43];;
      MELT_LOCATION ("warmelt-normal.melt:5705:/ putxtraresult");
      if (!meltxrestab_ || !meltxresdescr_)
	goto meltlabend_rout;
      if (meltxresdescr_[0] != MELTBPAR_PTR)
	goto meltlabend_rout;
      if (meltxrestab_[0].meltbp_aptr)
	*(meltxrestab_[0].meltbp_aptr) =
	  (melt_ptr_t) ( /*_.NREALBINDINGS__V46*/ meltfptr[42]);
      ;
      /*^finalreturn */
      ;
      /*finalret */ goto meltlabend_rout;

      {
	MELT_LOCATION ("warmelt-normal.melt:5706:/ locexp");
	/*void */ (void) 0;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;

      MELT_LOCATION ("warmelt-normal.melt:5700:/ clear");
	    /*clear *//*_.RETURN___V49*/ meltfptr[47] = 0;

      /*^clear */
	    /*clear *//*_.NREALBINDINGS__V46*/ meltfptr[42] = 0;

      MELT_LOCATION ("warmelt-normal.melt:5694:/ clear");
	    /*clear *//*_.TUPLE___V39*/ meltfptr[32] = 0;
      /*^clear */
	    /*clear *//*_.SPRIM__V41*/ meltfptr[33] = 0;

      MELT_LOCATION ("warmelt-normal.melt:5693:/ clear");
	    /*clear *//*_.S0__V38*/ meltfptr[35] = 0;
      /*^objgoto */
      /*objgoto */ goto mtch1__end /*endmatch */ ;
      ;

    /*objlabel */ mtch1_2:;
      MELT_LOCATION ("warmelt-normal.melt:5708:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_10;
      meltlab_count_10++;
      debugeprintf
	("objlabel_10 mtch1_*CLASS_NORMTESTER_TUPLE/1aef41d2. *CLASS_OBJLABELINSTR/1aadd390. %ld",
	 meltlab_count_10);
#endif

      ;
      /*^clear */
	    /*clear *//*_.NTHCOMP__V50*/ meltfptr[46] = 0;
      /*^clear */
	    /*clear *//*_.NTHCOMP__V51*/ meltfptr[47] = 0;
      /*^cond */
      /*cond */ if (
		     /*normtestertuple */
	(melt_magic_discr ((melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15])) ==
	 MELTOBMAG_MULTIPLE
	 && ((meltmultiple_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]))->nbval == 2))	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

	    /*^compute */
    /*_.NTHCOMP__V50*/ meltfptr[46] =
	      /*unsafenthcomp */
	      ((meltmultiple_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]))->
	      tabval[0];;
	    /*^compute */
    /*_.NTHCOMP__V51*/ meltfptr[47] =
	      /*unsafenthcomp */
	      ((meltmultiple_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]))->
	      tabval[1];;
	    /*^objgoto */
	    /*objgoto */ goto mtch1_3;
	    ;
	  }
	  ;
	}
      else
	{			/*^cond.else */

	  /*^block */
	  /*anyblock */
	  {

	    MELT_LOCATION ("warmelt-normal.melt:5723:/ objgoto");
	    /*objgoto */ goto mtch1_4;
	    ;
	  }
	  ;
	}
      ;

    /*objlabel */ mtch1_3:;
      MELT_LOCATION ("warmelt-normal.melt:5708:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_11;
      meltlab_count_11++;
      debugeprintf
	("objlabel_11 mtch1_*CLASS_NORMTESTER_SUCCESS/2cf06a18. *CLASS_OBJLABELINSTR/1777dcf6. %ld",
	 meltlab_count_11);
#endif

      ;
      /*^quasiblock */


      /*_.S0__V52*/ meltfptr[43] = /*_.NTHCOMP__V50*/ meltfptr[46];;
      /*^compute */
      /*_.S1__V53*/ meltfptr[42] = /*_.NTHCOMP__V51*/ meltfptr[47];;
      MELT_LOCATION ("warmelt-normal.melt:5709:/ quasiblock");



      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5712:/ blockmultialloc");
      /*multiallocblock */
      {
	struct meltletrec_2_st
	{
	  struct MELT_MULTIPLE_STRUCT (2) rtup_0__TUPLREC__x13;
	  long meltletrec_2_endgap;
	} *meltletrec_2_ptr = 0;
	meltletrec_2_ptr =
	  (struct meltletrec_2_st *)
	  meltgc_allocate (sizeof (struct meltletrec_2_st), 0);
	/*^blockmultialloc.initfill */
	/*inimult rtup_0__TUPLREC__x13 */
 /*_.TUPLREC___V55*/ meltfptr[33] =
	  (melt_ptr_t) & meltletrec_2_ptr->rtup_0__TUPLREC__x13;
	meltletrec_2_ptr->rtup_0__TUPLREC__x13.discr =
	  (meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_MULTIPLE))));
	meltletrec_2_ptr->rtup_0__TUPLREC__x13.nbval = 2;


	/*^putuple */
	/*putupl#43 */
	melt_assertmsg ("putupl [:5712] #43 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V55*/ meltfptr[33]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5712] #43 checkoff",
			(0 >= 0
			 && 0 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V55*/
						meltfptr[33]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V55*/ meltfptr[33]))->
	  tabval[0] = (melt_ptr_t) ( /*_.S0__V52*/ meltfptr[43]);
	;
	/*^putuple */
	/*putupl#44 */
	melt_assertmsg ("putupl [:5712] #44 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V55*/ meltfptr[33]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5712] #44 checkoff",
			(1 >= 0
			 && 1 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V55*/
						meltfptr[33]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V55*/ meltfptr[33]))->
	  tabval[1] = (melt_ptr_t) ( /*_.S1__V53*/ meltfptr[42]);
	;
	/*^touch */
	meltgc_touch ( /*_.TUPLREC___V55*/ meltfptr[33]);
	;
	/*_.TUPLE___V54*/ meltfptr[32] = /*_.TUPLREC___V55*/ meltfptr[33];;
	/*epilog */

	MELT_LOCATION ("warmelt-normal.melt:5712:/ clear");
	     /*clear *//*_.TUPLREC___V55*/ meltfptr[33] = 0;
	/*^clear */
	     /*clear *//*_.TUPLREC___V55*/ meltfptr[33] = 0;
      }				/*end multiallocblock */
      ;
      MELT_LOCATION ("warmelt-normal.melt:5709:/ quasiblock");


      /*^rawallocobj */
      /*rawallocobj */
      {
	melt_ptr_t newobj = 0;
	melt_raw_object_create (newobj,
				(melt_ptr_t) (( /*!CLASS_SOURCE_PRIMITIVE */
					       meltfrout->tabval[8])), (4),
				"CLASS_SOURCE_PRIMITIVE");
   /*_.INST__V57*/ meltfptr[33] =
	  newobj;
      };
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LOCA_LOCATION",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V57*/ meltfptr[33])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V57*/ meltfptr[33]), (1),
			    ( /*_.SLOC__V15*/ meltfptr[14]), "LOCA_LOCATION");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SPRIM_OPER",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V57*/ meltfptr[33])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V57*/ meltfptr[33]), (3),
			    ( /*_.SPRIMITIVE__V18*/ meltfptr[17]),
			    "SPRIM_OPER");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SARGOP_ARGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V57*/ meltfptr[33])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V57*/ meltfptr[33]), (2),
			    ( /*_.TUPLE___V54*/ meltfptr[32]), "SARGOP_ARGS");
      ;
      /*^touchobj */

      melt_dbgtrace_written_object ( /*_.INST__V57*/ meltfptr[33],
				    "newly made instance");
      ;
      /*_.SPRIM__V56*/ meltfptr[35] = /*_.INST__V57*/ meltfptr[33];;

#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5714:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L19*/ meltfnum[0] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5714:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L19*/ meltfnum[0])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L20*/ meltfnum[15] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5714:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[5];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L20*/ meltfnum[15];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5714;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation binary sprim=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.SPRIM__V56*/ meltfptr[35];
		/*_.MELT_DEBUG_FUN__V59*/ meltfptr[58] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V58*/ meltfptr[57] =
		/*_.MELT_DEBUG_FUN__V59*/ meltfptr[58];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5714:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L20*/ meltfnum[15] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V59*/ meltfptr[58] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V58*/ meltfptr[57] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5714:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L19*/ meltfnum[0] = 0;
	/*^clear */
	      /*clear *//*_._IF___V58*/ meltfptr[57] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5715:/ quasiblock");


      /*^multiapply */
      /*multiapply 4args, 1x.res */
      {
	union meltparam_un argtab[3];

	union meltparam_un restab[1];
	memset (&restab, 0, sizeof (restab));
	memset (&argtab, 0, sizeof (argtab));
	/*^multiapply.arg */
	argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];	/*^multiapply.arg */
	argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
	argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.PSLOC__V5*/ meltfptr[4];
	/*^multiapply.xres */
	restab[0].meltbp_aptr =
	  (melt_ptr_t *) & /*_.NREALBINDINGS__V61*/ meltfptr[57];
	/*^multiapply.appl */
	/*_.NREALARGS__V60*/ meltfptr[58] =
	  melt_apply ((meltclosure_ptr_t)
		      (( /*!NORMEXP_PRIMITIVE */ meltfrout->tabval[9])),
		      (melt_ptr_t) ( /*_.SPRIM__V56*/ meltfptr[35]),
		      (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		      argtab, (MELTBPARSTR_PTR ""), restab);
      }
      ;
      /*^quasiblock */



#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5718:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L21*/ meltfnum[15] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5718:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L21*/ meltfnum[15])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L22*/ meltfnum[0] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5718:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[7];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L22*/ meltfnum[0];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5718;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation binary result nrealargs=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALARGS__V60*/ meltfptr[58];
		/*^apply.arg */
		argtab[5].meltbp_cstring = " nrealbindings=";
		/*^apply.arg */
		argtab[6].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALBINDINGS__V61*/ meltfptr[57];
		/*_.MELT_DEBUG_FUN__V63*/ meltfptr[62] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V62*/ meltfptr[61] =
		/*_.MELT_DEBUG_FUN__V63*/ meltfptr[62];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5718:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L22*/ meltfnum[0] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V63*/ meltfptr[62] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V62*/ meltfptr[61] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5718:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L21*/ meltfnum[15] = 0;
	/*^clear */
	      /*clear *//*_._IF___V62*/ meltfptr[61] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5720:/ quasiblock");


      /*_._RETVAL___V1*/ meltfptr[0] = /*_.NREALARGS__V60*/ meltfptr[58];;
      MELT_LOCATION ("warmelt-normal.melt:5720:/ putxtraresult");
      if (!meltxrestab_ || !meltxresdescr_)
	goto meltlabend_rout;
      if (meltxresdescr_[0] != MELTBPAR_PTR)
	goto meltlabend_rout;
      if (meltxrestab_[0].meltbp_aptr)
	*(meltxrestab_[0].meltbp_aptr) =
	  (melt_ptr_t) ( /*_.NREALBINDINGS__V61*/ meltfptr[57]);
      ;
      /*^finalreturn */
      ;
      /*finalret */ goto meltlabend_rout;

      {
	MELT_LOCATION ("warmelt-normal.melt:5721:/ locexp");
	/*void */ (void) 0;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;

      MELT_LOCATION ("warmelt-normal.melt:5715:/ clear");
	    /*clear *//*_.RETURN___V64*/ meltfptr[62] = 0;

      /*^clear */
	    /*clear *//*_.NREALBINDINGS__V61*/ meltfptr[57] = 0;

      MELT_LOCATION ("warmelt-normal.melt:5709:/ clear");
	    /*clear *//*_.TUPLE___V54*/ meltfptr[32] = 0;
      /*^clear */
	    /*clear *//*_.SPRIM__V56*/ meltfptr[35] = 0;

      MELT_LOCATION ("warmelt-normal.melt:5708:/ clear");
	    /*clear *//*_.S0__V52*/ meltfptr[43] = 0;
      /*^clear */
	    /*clear *//*_.S1__V53*/ meltfptr[42] = 0;
      /*^objgoto */
      /*objgoto */ goto mtch1__end /*endmatch */ ;
      ;

    /*objlabel */ mtch1_4:;
      MELT_LOCATION ("warmelt-normal.melt:5723:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_12;
      meltlab_count_12++;
      debugeprintf
	("objlabel_12 mtch1_*CLASS_NORMTESTER_SUCCESS/3392e7ed. *CLASS_OBJLABELINSTR/7d92465. %ld",
	 meltlab_count_12);
#endif

      ;
      /*^quasiblock */



#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5724:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	MELT_CHECK_SIGNAL ();
	;
    /*_#MULTIPLE_LENGTH__L23*/ meltfnum[0] =
	  (melt_multiple_length
	   ((melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15])));;
	/*^compute */
    /*_#gtI__L24*/ meltfnum[15] =
	  (( /*_#MULTIPLE_LENGTH__L23*/ meltfnum[0]) > (2));;
	MELT_LOCATION ("warmelt-normal.melt:5724:/ cond");
	/*cond */ if ( /*_#gtI__L24*/ meltfnum[15])	/*then */
	  {
	    /*^cond.then */
	    /*_._IFELSE___V66*/ meltfptr[62] = ( /*nil */ NULL);;
	  }
	else
	  {
	    MELT_LOCATION ("warmelt-normal.melt:5724:/ cond.else");

	    /*^block */
	    /*anyblock */
	    {




	      {
		/*^locexp */
		melt_assert_failed (("check #args>2"),
				    ("warmelt-normal.melt")
				    ? ("warmelt-normal.melt") : __FILE__,
				    (5724) ? (5724) : __LINE__, __FUNCTION__);
		;
	      }
	      ;
		/*clear *//*_._IFELSE___V66*/ meltfptr[62] = 0;
	      /*epilog */
	    }
	    ;
	  }
	;
	/*^compute */
	/*_.IFCPP___V65*/ meltfptr[61] = /*_._IFELSE___V66*/ meltfptr[62];;
	/*epilog */

	MELT_LOCATION ("warmelt-normal.melt:5724:/ clear");
	      /*clear *//*_#MULTIPLE_LENGTH__L23*/ meltfnum[0] = 0;
	/*^clear */
	      /*clear *//*_#gtI__L24*/ meltfnum[15] = 0;
	/*^clear */
	      /*clear *//*_._IFELSE___V66*/ meltfptr[62] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*_.IFCPP___V65*/ meltfptr[61] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
      ;
      MELT_LOCATION ("warmelt-normal.melt:5725:/ quasiblock");



      MELT_CHECK_SIGNAL ();
      ;
  /*_.MULTIPLE_NTH__V67*/ meltfptr[58] =
	(melt_multiple_nth
	 ((melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]), (0)));;
      /*^compute */
  /*_.MULTIPLE_NTH__V68*/ meltfptr[57] =
	(melt_multiple_nth
	 ((melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]), (1)));;
      MELT_LOCATION ("warmelt-normal.melt:5728:/ blockmultialloc");
      /*multiallocblock */
      {
	struct meltletrec_3_st
	{
	  struct MELT_MULTIPLE_STRUCT (2) rtup_0__TUPLREC__x14;
	  long meltletrec_3_endgap;
	} *meltletrec_3_ptr = 0;
	meltletrec_3_ptr =
	  (struct meltletrec_3_st *)
	  meltgc_allocate (sizeof (struct meltletrec_3_st), 0);
	/*^blockmultialloc.initfill */
	/*inimult rtup_0__TUPLREC__x14 */
 /*_.TUPLREC___V70*/ meltfptr[35] =
	  (melt_ptr_t) & meltletrec_3_ptr->rtup_0__TUPLREC__x14;
	meltletrec_3_ptr->rtup_0__TUPLREC__x14.discr =
	  (meltobject_ptr_t) (((melt_ptr_t) (MELT_PREDEF (DISCR_MULTIPLE))));
	meltletrec_3_ptr->rtup_0__TUPLREC__x14.nbval = 2;


	/*^putuple */
	/*putupl#45 */
	melt_assertmsg ("putupl [:5728] #45 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V70*/ meltfptr[35]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5728] #45 checkoff",
			(0 >= 0
			 && 0 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V70*/
						meltfptr[35]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V70*/ meltfptr[35]))->
	  tabval[0] = (melt_ptr_t) ( /*_.MULTIPLE_NTH__V67*/ meltfptr[58]);
	;
	/*^putuple */
	/*putupl#46 */
	melt_assertmsg ("putupl [:5728] #46 checktup",
			melt_magic_discr ((melt_ptr_t)
					  ( /*_.TUPLREC___V70*/ meltfptr[35]))
			== MELTOBMAG_MULTIPLE);
	melt_assertmsg ("putupl [:5728] #46 checkoff",
			(1 >= 0
			 && 1 <
			 melt_multiple_length ((melt_ptr_t)
					       ( /*_.TUPLREC___V70*/
						meltfptr[35]))));
	((meltmultiple_ptr_t) ( /*_.TUPLREC___V70*/ meltfptr[35]))->
	  tabval[1] = (melt_ptr_t) ( /*_.MULTIPLE_NTH__V68*/ meltfptr[57]);
	;
	/*^touch */
	meltgc_touch ( /*_.TUPLREC___V70*/ meltfptr[35]);
	;
	/*_.TUPLE___V69*/ meltfptr[32] = /*_.TUPLREC___V70*/ meltfptr[35];;
	/*epilog */

	MELT_LOCATION ("warmelt-normal.melt:5728:/ clear");
	     /*clear *//*_.TUPLREC___V70*/ meltfptr[35] = 0;
	/*^clear */
	     /*clear *//*_.TUPLREC___V70*/ meltfptr[35] = 0;
      }				/*end multiallocblock */
      ;
      MELT_LOCATION ("warmelt-normal.melt:5725:/ quasiblock");


      /*^rawallocobj */
      /*rawallocobj */
      {
	melt_ptr_t newobj = 0;
	melt_raw_object_create (newobj,
				(melt_ptr_t) (( /*!CLASS_SOURCE_PRIMITIVE */
					       meltfrout->tabval[8])), (4),
				"CLASS_SOURCE_PRIMITIVE");
   /*_.INST__V72*/ meltfptr[42] =
	  newobj;
      };
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @LOCA_LOCATION",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V72*/ meltfptr[42])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V72*/ meltfptr[42]), (1),
			    ( /*_.SLOC__V15*/ meltfptr[14]), "LOCA_LOCATION");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SPRIM_OPER",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V72*/ meltfptr[42])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V72*/ meltfptr[42]), (3),
			    ( /*_.SPRIMITIVE__V18*/ meltfptr[17]),
			    "SPRIM_OPER");
      ;
      /*^putslot */
      /*putslot */
      melt_assertmsg ("checkobj putslot  _ @SARGOP_ARGS",
		      melt_magic_discr ((melt_ptr_t)
					( /*_.INST__V72*/ meltfptr[42])) ==
		      MELTOBMAG_OBJECT);
      melt_putfield_object (( /*_.INST__V72*/ meltfptr[42]), (2),
			    ( /*_.TUPLE___V69*/ meltfptr[32]), "SARGOP_ARGS");
      ;
      /*^touchobj */

      melt_dbgtrace_written_object ( /*_.INST__V72*/ meltfptr[42],
				    "newly made instance");
      ;
      /*_.SPRIM__V71*/ meltfptr[43] = /*_.INST__V72*/ meltfptr[42];;
      /*citerblock FOREACH_IN_MULTIPLE */
      {
	/* start foreach_in_multiple meltcit2__EACHTUP */
	long meltcit2__EACHTUP_ln =
	  melt_multiple_length ((melt_ptr_t) /*_.SARGS__V16*/ meltfptr[15]);
	for ( /*_#SIX__L25*/ meltfnum[0] = 0;
	     ( /*_#SIX__L25*/ meltfnum[0] >= 0)
	     && ( /*_#SIX__L25*/ meltfnum[0] < meltcit2__EACHTUP_ln);
	/*_#SIX__L25*/ meltfnum[0]++)
	  {
	    /*_.CURSARG__V73*/ meltfptr[62] =
	      melt_multiple_nth ((melt_ptr_t)
				 ( /*_.SARGS__V16*/ meltfptr[15]),
				 /*_#SIX__L25*/ meltfnum[0]);




	    MELT_CHECK_SIGNAL ();
	    ;
   /*_#gtI__L26*/ meltfnum[15] =
	      (( /*_#SIX__L25*/ meltfnum[0]) > (1));;
	    MELT_LOCATION ("warmelt-normal.melt:5733:/ cond");
	    /*cond */ if ( /*_#gtI__L26*/ meltfnum[15])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

		  MELT_LOCATION ("warmelt-normal.melt:5734:/ quasiblock");



		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION
		    ("warmelt-normal.melt:5737:/ blockmultialloc");
		  /*multiallocblock */
		  {
		    struct meltletrec_4_st
		    {
		      struct MELT_MULTIPLE_STRUCT (2) rtup_0__TUPLREC__x15;
		      long meltletrec_4_endgap;
		    } *meltletrec_4_ptr = 0;
		    meltletrec_4_ptr =
		      (struct meltletrec_4_st *)
		      meltgc_allocate (sizeof (struct meltletrec_4_st), 0);
		    /*^blockmultialloc.initfill */
		    /*inimult rtup_0__TUPLREC__x15 */
 /*_.TUPLREC___V77*/ meltfptr[76] =
		      (melt_ptr_t) & meltletrec_4_ptr->rtup_0__TUPLREC__x15;
		    meltletrec_4_ptr->rtup_0__TUPLREC__x15.discr =
		      (meltobject_ptr_t) (((melt_ptr_t)
					   (MELT_PREDEF (DISCR_MULTIPLE))));
		    meltletrec_4_ptr->rtup_0__TUPLREC__x15.nbval = 2;


		    /*^putuple */
		    /*putupl#47 */
		    melt_assertmsg ("putupl [:5737] #47 checktup",
				    melt_magic_discr ((melt_ptr_t)
						      ( /*_.TUPLREC___V77*/
						       meltfptr[76])) ==
				    MELTOBMAG_MULTIPLE);
		    melt_assertmsg ("putupl [:5737] #47 checkoff",
				    (0 >= 0
				     && 0 <
				     melt_multiple_length ((melt_ptr_t)
							   ( /*_.TUPLREC___V77*/ meltfptr[76]))));
		    ((meltmultiple_ptr_t)
		     ( /*_.TUPLREC___V77*/ meltfptr[76]))->tabval[0] =
 (melt_ptr_t) ( /*_.SPRIM__V71*/ meltfptr[43]);
		    ;
		    /*^putuple */
		    /*putupl#48 */
		    melt_assertmsg ("putupl [:5737] #48 checktup",
				    melt_magic_discr ((melt_ptr_t)
						      ( /*_.TUPLREC___V77*/
						       meltfptr[76])) ==
				    MELTOBMAG_MULTIPLE);
		    melt_assertmsg ("putupl [:5737] #48 checkoff",
				    (1 >= 0
				     && 1 <
				     melt_multiple_length ((melt_ptr_t)
							   ( /*_.TUPLREC___V77*/ meltfptr[76]))));
		    ((meltmultiple_ptr_t)
		     ( /*_.TUPLREC___V77*/ meltfptr[76]))->tabval[1] =
 (melt_ptr_t) ( /*_.CURSARG__V73*/ meltfptr[62]);
		    ;
		    /*^touch */
		    meltgc_touch ( /*_.TUPLREC___V77*/ meltfptr[76]);
		    ;
		    /*_.TUPLE___V76*/ meltfptr[75] =
		      /*_.TUPLREC___V77*/ meltfptr[76];;
		    /*epilog */

		    MELT_LOCATION ("warmelt-normal.melt:5737:/ clear");
		/*clear *//*_.TUPLREC___V77*/ meltfptr[76] = 0;
		    /*^clear */
		/*clear *//*_.TUPLREC___V77*/ meltfptr[76] = 0;
		  }		/*end multiallocblock */
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5734:/ quasiblock");


		  /*^rawallocobj */
		  /*rawallocobj */
		  {
		    melt_ptr_t newobj = 0;
		    melt_raw_object_create (newobj,
					    (melt_ptr_t) (( /*!CLASS_SOURCE_PRIMITIVE */ meltfrout->tabval[8])), (4), "CLASS_SOURCE_PRIMITIVE");
      /*_.INST__V79*/ meltfptr[78] =
		      newobj;
		  };
		  ;
		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @LOCA_LOCATION",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.INST__V79*/
						     meltfptr[78])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.INST__V79*/ meltfptr[78]), (1),
					( /*_.SLOC__V15*/ meltfptr[14]),
					"LOCA_LOCATION");
		  ;
		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @SPRIM_OPER",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.INST__V79*/
						     meltfptr[78])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.INST__V79*/ meltfptr[78]), (3),
					( /*_.SPRIMITIVE__V18*/ meltfptr[17]),
					"SPRIM_OPER");
		  ;
		  /*^putslot */
		  /*putslot */
		  melt_assertmsg ("checkobj putslot  _ @SARGOP_ARGS",
				  melt_magic_discr ((melt_ptr_t)
						    ( /*_.INST__V79*/
						     meltfptr[78])) ==
				  MELTOBMAG_OBJECT);
		  melt_putfield_object (( /*_.INST__V79*/ meltfptr[78]), (2),
					( /*_.TUPLE___V76*/ meltfptr[75]),
					"SARGOP_ARGS");
		  ;
		  /*^touchobj */

		  melt_dbgtrace_written_object ( /*_.INST__V79*/ meltfptr[78],
						"newly made instance");
		  ;
		  /*_.NEWSPRIM__V78*/ meltfptr[76] =
		    /*_.INST__V79*/ meltfptr[78];;
		  MELT_LOCATION ("warmelt-normal.melt:5739:/ compute");
		  /*_.SPRIM__V71*/ meltfptr[43] =
		    /*_.SETQ___V80*/ meltfptr[79] =
		    /*_.NEWSPRIM__V78*/ meltfptr[76];;
		  /*_.LET___V75*/ meltfptr[74] =
		    /*_.SETQ___V80*/ meltfptr[79];;

		  MELT_LOCATION ("warmelt-normal.melt:5734:/ clear");
	       /*clear *//*_.TUPLE___V76*/ meltfptr[75] = 0;
		  /*^clear */
	       /*clear *//*_.NEWSPRIM__V78*/ meltfptr[76] = 0;
		  /*^clear */
	       /*clear *//*_.SETQ___V80*/ meltfptr[79] = 0;
		  /*_._IF___V74*/ meltfptr[35] =
		    /*_.LET___V75*/ meltfptr[74];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5733:/ clear");
	       /*clear *//*_.LET___V75*/ meltfptr[74] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

    /*_._IF___V74*/ meltfptr[35] = NULL;;
	      }
	    ;
	    if ( /*_#SIX__L25*/ meltfnum[0] < 0)
	      break;
	  }			/* end  foreach_in_multiple meltcit2__EACHTUP */

	/*citerepilog */

	MELT_LOCATION ("warmelt-normal.melt:5730:/ clear");
	     /*clear *//*_.CURSARG__V73*/ meltfptr[62] = 0;
	/*^clear */
	     /*clear *//*_#SIX__L25*/ meltfnum[0] = 0;
	/*^clear */
	     /*clear *//*_#gtI__L26*/ meltfnum[15] = 0;
	/*^clear */
	     /*clear *//*_._IF___V74*/ meltfptr[35] = 0;
      }				/*endciterblock FOREACH_IN_MULTIPLE */
      ;

#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5741:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L27*/ meltfnum[26] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5741:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L27*/ meltfnum[26])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L28*/ meltfnum[27] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5741:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[5];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L28*/ meltfnum[27];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5741;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation nary sprim=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.SPRIM__V71*/ meltfptr[43];
		/*_.MELT_DEBUG_FUN__V82*/ meltfptr[76] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V81*/ meltfptr[75] =
		/*_.MELT_DEBUG_FUN__V82*/ meltfptr[76];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5741:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L28*/ meltfnum[27] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V82*/ meltfptr[76] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V81*/ meltfptr[75] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5741:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L27*/ meltfnum[26] = 0;
	/*^clear */
	      /*clear *//*_._IF___V81*/ meltfptr[75] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5742:/ quasiblock");


      /*^multiapply */
      /*multiapply 4args, 1x.res */
      {
	union meltparam_un argtab[3];

	union meltparam_un restab[1];
	memset (&restab, 0, sizeof (restab));
	memset (&argtab, 0, sizeof (argtab));
	/*^multiapply.arg */
	argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];	/*^multiapply.arg */
	argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
	argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.PSLOC__V5*/ meltfptr[4];
	/*^multiapply.xres */
	restab[0].meltbp_aptr =
	  (melt_ptr_t *) & /*_.NREALBINDINGS__V84*/ meltfptr[74];
	/*^multiapply.appl */
	/*_.NREALARGS__V83*/ meltfptr[79] =
	  melt_apply ((meltclosure_ptr_t)
		      (( /*!NORMEXP_PRIMITIVE */ meltfrout->tabval[9])),
		      (melt_ptr_t) ( /*_.SPRIM__V71*/ meltfptr[43]),
		      (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		      argtab, (MELTBPARSTR_PTR ""), restab);
      }
      ;
      /*^quasiblock */



#if MELT_HAVE_DEBUG
      MELT_LOCATION ("warmelt-normal.melt:5745:/ cppif.then");
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	  melt_dbgcounter++;
#endif
	  ;
	}
	;

	MELT_CHECK_SIGNAL ();
	;
    /*_#MELT_NEED_DBG__L29*/ meltfnum[27] =
	  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	  0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	  ;;
	MELT_LOCATION ("warmelt-normal.melt:5745:/ cond");
	/*cond */ if ( /*_#MELT_NEED_DBG__L29*/ meltfnum[27])	/*then */
	  {
	    /*^cond.then */
	    /*^block */
	    /*anyblock */
	    {

      /*_#MELT_CALLCOUNT__L30*/ meltfnum[26] =
		/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		meltcallcount	/* melt_callcount debugging */
#else
		0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		;;

	      MELT_CHECK_SIGNAL ();
	      ;
	      MELT_LOCATION ("warmelt-normal.melt:5745:/ apply");
	      /*apply */
	      {
		union meltparam_un argtab[7];
		memset (&argtab, 0, sizeof (argtab));
		/*^apply.arg */
		argtab[0].meltbp_long =
		  /*_#MELT_CALLCOUNT__L30*/ meltfnum[26];
		/*^apply.arg */
		argtab[1].meltbp_cstring = "warmelt-normal.melt";
		/*^apply.arg */
		argtab[2].meltbp_long = 5745;
		/*^apply.arg */
		argtab[3].meltbp_cstring =
		  "normexp_arithmetic_variadic_operation nary result nrealargs=";
		/*^apply.arg */
		argtab[4].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALARGS__V83*/ meltfptr[79];
		/*^apply.arg */
		argtab[5].meltbp_cstring = " nrealbindings=";
		/*^apply.arg */
		argtab[6].meltbp_aptr =
		  (melt_ptr_t *) & /*_.NREALBINDINGS__V84*/ meltfptr[74];
		/*_.MELT_DEBUG_FUN__V86*/ meltfptr[75] =
		  melt_apply ((meltclosure_ptr_t)
			      (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			      (melt_ptr_t) (( /*nil */ NULL)),
			      (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			       MELTBPARSTR_PTR ""), argtab, "",
			      (union meltparam_un *) 0);
	      }
	      ;
	      /*_._IF___V85*/ meltfptr[76] =
		/*_.MELT_DEBUG_FUN__V86*/ meltfptr[75];;
	      /*epilog */

	      MELT_LOCATION ("warmelt-normal.melt:5745:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L30*/ meltfnum[26] = 0;
	      /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V86*/ meltfptr[75] = 0;
	    }
	    ;
	  }
	else
	  {			/*^cond.else */

     /*_._IF___V85*/ meltfptr[76] = NULL;;
	  }
	;

	{
	  MELT_LOCATION ("warmelt-normal.melt:5745:/ locexp");
	  /*void */ (void) 0;
	}
	;
	/*^quasiblock */


	/*epilog */

	/*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L29*/ meltfnum[27] = 0;
	/*^clear */
	      /*clear *//*_._IF___V85*/ meltfptr[76] = 0;
      }

#else /*MELT_HAVE_DEBUG */
      /*^cppif.else */
      /*^block */
      /*anyblock */
      {


	{
	  /*^locexp */
	  /*void */ (void) 0;
	}
	;
	/*epilog */
      }

#endif /*MELT_HAVE_DEBUG */
      ;

      MELT_CHECK_SIGNAL ();
      ;
      MELT_LOCATION ("warmelt-normal.melt:5747:/ quasiblock");


      /*_._RETVAL___V1*/ meltfptr[0] = /*_.NREALARGS__V83*/ meltfptr[79];;
      MELT_LOCATION ("warmelt-normal.melt:5747:/ putxtraresult");
      if (!meltxrestab_ || !meltxresdescr_)
	goto meltlabend_rout;
      if (meltxresdescr_[0] != MELTBPAR_PTR)
	goto meltlabend_rout;
      if (meltxrestab_[0].meltbp_aptr)
	*(meltxrestab_[0].meltbp_aptr) =
	  (melt_ptr_t) ( /*_.NREALBINDINGS__V84*/ meltfptr[74]);
      ;
      /*^finalreturn */
      ;
      /*finalret */ goto meltlabend_rout;

      {
	MELT_LOCATION ("warmelt-normal.melt:5748:/ locexp");
	/*void */ (void) 0;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;

      MELT_LOCATION ("warmelt-normal.melt:5742:/ clear");
	    /*clear *//*_.RETURN___V87*/ meltfptr[75] = 0;

      /*^clear */
	    /*clear *//*_.NREALBINDINGS__V84*/ meltfptr[74] = 0;

      MELT_LOCATION ("warmelt-normal.melt:5725:/ clear");
	    /*clear *//*_.MULTIPLE_NTH__V67*/ meltfptr[58] = 0;
      /*^clear */
	    /*clear *//*_.MULTIPLE_NTH__V68*/ meltfptr[57] = 0;
      /*^clear */
	    /*clear *//*_.TUPLE___V69*/ meltfptr[32] = 0;
      /*^clear */
	    /*clear *//*_.SPRIM__V71*/ meltfptr[43] = 0;

      MELT_CHECK_SIGNAL ();
      ;

      MELT_LOCATION ("warmelt-normal.melt:5723:/ clear");
	    /*clear *//*_.IFCPP___V65*/ meltfptr[61] = 0;
      /*^objgoto */
      /*objgoto */ goto mtch1__end /*endmatch */ ;
      ;

    /*objlabel */ mtch1__end:;
      MELT_LOCATION ("warmelt-normal.melt:5691:/ objlabel");
      /* objlabel */
#if MELTDEBUG_MATCHING
      static long meltlab_count_13;
      meltlab_count_13++;
      debugeprintf
	("objlabel_13 mtch1__end_h584011217 *CLASS_OBJLABELINSTR/22cf4dd1. %ld",
	 meltlab_count_13);
#endif

      ;
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;

    MELT_LOCATION ("warmelt-normal.melt:5676:/ clear");
	   /*clear *//*_.IFCPP___V29*/ meltfptr[27] = 0;

    /*^clear */
	   /*clear *//*_.NBINDINGS__V26*/ meltfptr[25] = 0;

    MELT_LOCATION ("warmelt-normal.melt:5666:/ clear");
	   /*clear *//*_.LOCA_LOCATION__V14*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.SLOC__V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.SARGS__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.SNEUTRAL__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.SPRIMITIVE__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V21*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V23*/ meltfptr[18] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */

    MELT_LOCATION ("warmelt-normal.melt:5661:/ clear");
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V12*/ meltfptr[10] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMEXP_ARITHMETIC_VARIADIC_OPERATION", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_101_WARMELTmiNORMAL_NORMEXP_ARITHMETIC_VARIADIC_OPERATION */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING
  (meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_,
   const melt_argdescr_cell_t meltxargdescr_[],
   union meltparam_un * meltxargtab_,
   const melt_argdescr_cell_t meltxresdescr_[],
   union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 12
    melt_ptr_t mcfr_varptr[12];
#define MELTFRAM_NBVARNUM 2
    long mcfr_varnum[2];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct
	 meltframe_meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING_st
	 *) meltfirstargp_;
      /* use arguments meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 12; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING nbval 12*/
  meltfram__.mcfr_nbvar = 12 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("BADMETH_PREPARE_CONSTRUCTOR_BINDING", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5762:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5763:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5763:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5763:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5763;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"bad_prepare_constructor_binding recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5763:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5763:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5764:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.SYMB__V3*/ meltfptr[2]),
					(melt_ptr_t) (( /*!CLASS_NAMED */
						       meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "NAMED_NAME");
   /*_.NAMED_NAME__V8*/ meltfptr[6] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NAMED_NAME__V8*/ meltfptr[6] = NULL;;
      }
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5764:/ locexp");
      melt_error_str ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4]),
		      ("invalid constructor binding"),
		      (melt_ptr_t) ( /*_.NAMED_NAME__V8*/ meltfptr[6]));
    }
    ;
 /*_.DISCRIM__V9*/ meltfptr[5] =
      ((melt_ptr_t)
       (melt_discr ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]))));;
    MELT_LOCATION ("warmelt-normal.melt:5766:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.DISCRIM__V9*/ meltfptr[5]),
					(melt_ptr_t) (( /*!CLASS_NAMED */
						       meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.DISCRIM__V9*/ meltfptr[5]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "NAMED_NAME");
   /*_.NAMED_NAME__V10*/ meltfptr[9] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NAMED_NAME__V10*/ meltfptr[9] = NULL;;
      }
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5765:/ locexp");
      melt_error_str ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4]),
		      ("receiver unprepared to constructor binding"),
		      (melt_ptr_t) ( /*_.NAMED_NAME__V10*/ meltfptr[9]));
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5767:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
      /*^cond */
      /*cond */ if (( /*nil */ NULL))	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V12*/ meltfptr[11] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5767:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("@$@badmeth_prepare_constructor_binding"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5767) ? (5767) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V12*/ meltfptr[11] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V11*/ meltfptr[10] = /*_._IFELSE___V12*/ meltfptr[11];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5767:/ clear");
	     /*clear *//*_._IFELSE___V12*/ meltfptr[11] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V11*/ meltfptr[10] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5762:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.IFCPP___V11*/ meltfptr[10];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5762:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.NAMED_NAME__V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.DISCRIM__V9*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.NAMED_NAME__V10*/ meltfptr[9] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V11*/ meltfptr[10] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("BADMETH_PREPARE_CONSTRUCTOR_BINDING", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_102_WARMELTmiNORMAL_BADMETH_PREPARE_CONSTRUCTOR_BINDING */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE
  (meltclosure_ptr_t meltclosp_, melt_ptr_t meltfirstargp_,
   const melt_argdescr_cell_t meltxargdescr_[],
   union meltparam_un * meltxargtab_,
   const melt_argdescr_cell_t meltxresdescr_[],
   union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 14
    melt_ptr_t mcfr_varptr[14];
#define MELTFRAM_NBVARNUM 2
    long mcfr_varnum[2];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct
	 meltframe_meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE_st
	 *) meltfirstargp_;
      /* use arguments meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 14; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE nbval 14*/
  meltfram__.mcfr_nbvar = 14 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("BADMETH_NORMAL_LETREC_CONSTRUCTIVE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5780:/ getarg");
 /*_.CEXPR__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.CBIND__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4])) != NULL);


  /*getarg#4 */
  /*^getarg */
  if (meltxargdescr_[3] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V6*/ meltfptr[5] =
    (meltxargtab_[3].meltbp_aptr) ? (*(meltxargtab_[3].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5])) != NULL);


  /*getarg#5 */
  /*^getarg */
  if (meltxargdescr_[4] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V7*/ meltfptr[6] =
    (meltxargtab_[4].meltbp_aptr) ? (*(meltxargtab_[4].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5781:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5781:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5781:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5781;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"badmeth_normal_letrec_constructive cexpr=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CEXPR__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " cbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V8*/ meltfptr[7] =
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5781:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V8*/ meltfptr[7] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5781:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5783:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.SYMB__V3*/ meltfptr[2]),
					(melt_ptr_t) (( /*!CLASS_NAMED */
						       meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "NAMED_NAME");
   /*_.NAMED_NAME__V10*/ meltfptr[8] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NAMED_NAME__V10*/ meltfptr[8] = NULL;;
      }
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5783:/ locexp");
      melt_error_str ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6]),
		      ("invalid letrec constructive symbol"),
		      (melt_ptr_t) ( /*_.NAMED_NAME__V10*/ meltfptr[8]));
    }
    ;
 /*_.DISCRIM__V11*/ meltfptr[7] =
      ((melt_ptr_t)
       (melt_discr ((melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]))));;
    MELT_LOCATION ("warmelt-normal.melt:5785:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.DISCRIM__V11*/ meltfptr[7]),
					(melt_ptr_t) (( /*!CLASS_NAMED */
						       meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.DISCRIM__V11*/ meltfptr[7]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "NAMED_NAME");
   /*_.NAMED_NAME__V12*/ meltfptr[11] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NAMED_NAME__V12*/ meltfptr[11] = NULL;;
      }
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5784:/ locexp");
      melt_error_str ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6]),
		      ("receiver unprepared to normal_letrec_constructive"),
		      (melt_ptr_t) ( /*_.NAMED_NAME__V12*/ meltfptr[11]));
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5786:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
      /*^cond */
      /*cond */ if (( /*nil */ NULL))	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V14*/ meltfptr[13] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5786:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("@$@badmeth_normal_letrec_constructive"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5786) ? (5786) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V14*/ meltfptr[13] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V13*/ meltfptr[12] = /*_._IFELSE___V14*/ meltfptr[13];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5786:/ clear");
	     /*clear *//*_._IFELSE___V14*/ meltfptr[13] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V13*/ meltfptr[12] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5780:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.IFCPP___V13*/ meltfptr[12];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5780:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.NAMED_NAME__V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.DISCRIM__V11*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_.NAMED_NAME__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V13*/ meltfptr[12] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("BADMETH_NORMAL_LETREC_CONSTRUCTIVE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_103_WARMELTmiNORMAL_BADMETH_NORMAL_LETREC_CONSTRUCTIVE */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un *
					      meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un *
					      meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 16
    melt_ptr_t mcfr_varptr[16];
#define MELTFRAM_NBVARNUM 5
    long mcfr_varnum[5];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA_st *)
	meltfirstargp_;
      /* use arguments meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 16; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA nbval 16*/
  meltfram__.mcfr_nbvar = 16 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("PREPCONS_LAMBDA", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5791:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5792:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5792:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5792:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5792;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_lambda recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5792:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5792:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5793:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_LAMBDA */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5793:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5793:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5793) ? (5793) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5793:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5794:/ quasiblock");



    MELT_CHECK_SIGNAL ();
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5797:/ apply");
    /*apply */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^apply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V5*/ meltfptr[4];
      /*^apply.arg */
      argtab[2].meltbp_cstring = "discr_closure";
      /*_.NORMAL_PREDEF__V11*/ meltfptr[10] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMAL_PREDEF */ meltfrout->tabval[3])),
		    (melt_ptr_t) (( /*!DISCR_CLOSURE */ meltfrout->
				   tabval[4])),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING ""),
		    argtab, "", (union meltparam_un *) 0);
    }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5794:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LAMBDA_BINDING */ meltfrout->tabval[2])), (7), "CLASS_NORMAL_CONSTRUCTED_LAMBDA_BINDING");
  /*_.INST__V13*/ meltfptr[12] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @BINDER",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V13*/ meltfptr[12])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V13*/ meltfptr[12]), (0),
			  ( /*_.SYMB__V3*/ meltfptr[2]), "BINDER");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V13*/ meltfptr[12])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V13*/ meltfptr[12]), (1),
			  ( /*_.SLOC__V5*/ meltfptr[4]), "NCONSB_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V13*/ meltfptr[12])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V13*/ meltfptr[12]), (2),
			  ( /*_.NORMAL_PREDEF__V11*/ meltfptr[10]),
			  "NCONSB_DISCR");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V13*/ meltfptr[12],
				  "newly made instance");
    ;
    /*_.CONSLAM__V12*/ meltfptr[11] = /*_.INST__V13*/ meltfptr[12];;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5800:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5800:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5800:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5800;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_lambda gives conslam";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSLAM__V12*/ meltfptr[11];
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V14*/ meltfptr[13] =
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5800:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V14*/ meltfptr[13] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5800:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V14*/ meltfptr[13] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5801:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.CONSLAM__V12*/ meltfptr[11];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5801:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V10*/ meltfptr[5] = /*_.RETURN___V16*/ meltfptr[14];;

    MELT_LOCATION ("warmelt-normal.melt:5794:/ clear");
	   /*clear *//*_.NORMAL_PREDEF__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.CONSLAM__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V16*/ meltfptr[14] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5791:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V10*/ meltfptr[5];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5791:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.LET___V10*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("PREPCONS_LAMBDA", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_104_WARMELTmiNORMAL_PREPCONS_LAMBDA */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un *
						meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un *
						meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 29
    melt_ptr_t mcfr_varptr[29];
#define MELTFRAM_NBVARNUM 11
    long mcfr_varnum[11];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA_st *)
	meltfirstargp_;
      /* use arguments meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 29; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA nbval 29*/
  meltfram__.mcfr_nbvar = 29 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMLETREC_LAMBDA", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5808:/ getarg");
 /*_.CEXPR__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.CBIND__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4])) != NULL);


  /*getarg#4 */
  /*^getarg */
  if (meltxargdescr_[3] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V6*/ meltfptr[5] =
    (meltxargtab_[3].meltbp_aptr) ? (*(meltxargtab_[3].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5])) != NULL);


  /*getarg#5 */
  /*^getarg */
  if (meltxargdescr_[4] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.PSLOC__V7*/ meltfptr[6] =
    (meltxargtab_[4].meltbp_aptr) ? (*(meltxargtab_[4].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.PSLOC__V7*/ meltfptr[6])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5809:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5809:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5809:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[11];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5809;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_lambda cexpr=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CEXPR__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " cbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*^apply.arg */
	      argtab[9].meltbp_cstring = " env=";
	      /*^apply.arg */
	      argtab[10].meltbp_aptr =
		(melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V8*/ meltfptr[7] =
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5809:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V8*/ meltfptr[7] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5809:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5811:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LAMBDA_BINDING */ meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5811:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[7] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5811:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cbind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5811) ? (5811) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[8] = /*_._IFELSE___V11*/ meltfptr[7];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5811:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[8] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5812:/ quasiblock");



    MELT_CHECK_SIGNAL ();
    ;
    /*^apply */
    /*apply */
    {
      /*_.NEWENV__V12*/ meltfptr[7] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!FRESH_ENV */ meltfrout->tabval[2])),
		    (melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4]), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5814:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5814:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5814:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5814;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_lambda newenv";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NEWENV__V12*/ meltfptr[7];
	      /*_.MELT_DEBUG_FUN__V14*/ meltfptr[13] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V13*/ meltfptr[12] =
	      /*_.MELT_DEBUG_FUN__V14*/ meltfptr[13];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5814:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V14*/ meltfptr[13] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V13*/ meltfptr[12] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5814:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V13*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5815:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L6*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NEWENV__V12*/ meltfptr[7]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[3])));;
      MELT_LOCATION ("warmelt-normal.melt:5815:/ cond");
      /*cond */ if ( /*_#IS_A__L6*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V16*/ meltfptr[12] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5815:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check newenv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5815) ? (5815) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V16*/ meltfptr[12] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V15*/ meltfptr[13] = /*_._IFELSE___V16*/ meltfptr[12];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5815:/ clear");
	     /*clear *//*_#IS_A__L6*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V16*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V15*/ meltfptr[13] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5816:/ quasiblock");


    /*^multiapply */
    /*multiapply 5args, 8x.res */
    {
      union meltparam_un argtab[4];

      union meltparam_un restab[8];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NEWENV__V12*/ meltfptr[7];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V6*/ meltfptr[5];	/*^multiapply.arg */
      argtab[3].meltbp_aptr = (melt_ptr_t *) & /*_.PSLOC__V7*/ meltfptr[6];
      /*^multiapply.xres */
      restab[0].meltbp_aptr = (melt_ptr_t *) & /*_.CSYM__V18*/ meltfptr[17];
      /*^multiapply.xres */
      restab[1].meltbp_aptr = (melt_ptr_t *) & /*_.CLOCC__V19*/ meltfptr[18];
      /*^multiapply.xres */
      restab[2].meltbp_aptr =
	(melt_ptr_t *) & /*_.CONSTROUT__V20*/ meltfptr[19];
      /*^multiapply.xres */
      restab[3].meltbp_aptr =
	(melt_ptr_t *) & /*_.CLOVTUP__V21*/ meltfptr[20];
      /*^multiapply.xres */
      restab[4].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V22*/ meltfptr[21];
      /*^multiapply.xres */
      restab[5].meltbp_aptr =
	(melt_ptr_t *) & /*_.OLDPROC__V23*/ meltfptr[22];
      /*^multiapply.xres */
      restab[6].meltbp_aptr =
	(melt_ptr_t *) & /*_.NDATAROUT__V24*/ meltfptr[23];
      /*^multiapply.xres */
      restab[7].meltbp_longptr = & /*_#INSIDEFLAG__L7*/ meltfnum[0];
      /*^multiapply.appl */
      /*_.NPROC__V17*/ meltfptr[12] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_LAMBDA */ meltfrout->tabval[4])),
		    (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
		     MELTBPARSTR_PTR ""), argtab,
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
		     MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
		     MELTBPARSTR_PTR MELTBPARSTR_LONG ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5820:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L8*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5820:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L8*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L9*/ meltfnum[8] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5820:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[15];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L9*/ meltfnum[8];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5820;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normletrec_lambda after normalize_lambda csym=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CSYM__V18*/ meltfptr[17];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " clocc=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.CLOCC__V19*/ meltfptr[18];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " constrout=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSTROUT__V20*/ meltfptr[19];
	      /*^apply.arg */
	      argtab[9].meltbp_cstring = " clovtup=";
	      /*^apply.arg */
	      argtab[10].meltbp_aptr =
		(melt_ptr_t *) & /*_.CLOVTUP__V21*/ meltfptr[20];
	      /*^apply.arg */
	      argtab[11].meltbp_cstring = " oldproc=";
	      /*^apply.arg */
	      argtab[12].meltbp_aptr =
		(melt_ptr_t *) & /*_.OLDPROC__V23*/ meltfptr[22];
	      /*^apply.arg */
	      argtab[13].meltbp_cstring = " ndatarout=";
	      /*^apply.arg */
	      argtab[14].meltbp_aptr =
		(melt_ptr_t *) & /*_.NDATAROUT__V24*/ meltfptr[23];
	      /*_.MELT_DEBUG_FUN__V26*/ meltfptr[25] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V25*/ meltfptr[24] =
	      /*_.MELT_DEBUG_FUN__V26*/ meltfptr[25];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5820:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L9*/ meltfnum[8] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V26*/ meltfptr[25] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V25*/ meltfptr[24] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5820:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L8*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V25*/ meltfptr[24] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5826:/ cond");
    /*cond */ if ( /*_#INSIDEFLAG__L7*/ meltfnum[0])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:5827:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.OLDPROC__V23*/ meltfptr[22]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 5, "NRPRO_CONST");
    /*_.NRPRO_CONST__V27*/ meltfptr[25] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    meltgc_append_list ((melt_ptr_t)
				( /*_.NRPRO_CONST__V27*/ meltfptr[25]),
				(melt_ptr_t) ( /*_.NDATAROUT__V24*/
					      meltfptr[23]));
	  }
	  ;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:5826:/ clear");
	     /*clear *//*_.NRPRO_CONST__V27*/ meltfptr[25] = 0;
	}
	;
      }				/*noelse */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5828:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CBIND__V4*/ meltfptr[3]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LAMBDA_BINDING */ meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NLAMBDAB_NCLOSED",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.CBIND__V4*/ meltfptr[3])) ==
			  MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.CBIND__V4*/ meltfptr[3]), (4),
				( /*_.CLOVTUP__V21*/ meltfptr[20]),
				"NLAMBDAB_NCLOSED");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NLAMBDAB_CONSTROUT",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.CBIND__V4*/ meltfptr[3])) ==
			  MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.CBIND__V4*/ meltfptr[3]), (5),
				( /*_.CONSTROUT__V20*/ meltfptr[19]),
				"NLAMBDAB_CONSTROUT");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NLAMBDAB_DATAROUT",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.CBIND__V4*/ meltfptr[3])) ==
			  MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.CBIND__V4*/ meltfptr[3]), (6),
				( /*_.NDATAROUT__V24*/ meltfptr[23]),
				"NLAMBDAB_DATAROUT");
	  ;
	  /*^touch */
	  meltgc_touch ( /*_.CBIND__V4*/ meltfptr[3]);
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.CBIND__V4*/ meltfptr[3],
					"put-fields");
	  ;
	  /*epilog */
	}
	;
      }				/*noelse */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5832:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L10*/ meltfnum[8] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5832:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[8])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L11*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5832:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L11*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5832;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_lambda updated cbind";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*_.MELT_DEBUG_FUN__V29*/ meltfptr[25] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V28*/ meltfptr[24] =
	      /*_.MELT_DEBUG_FUN__V29*/ meltfptr[25];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5832:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V29*/ meltfptr[25] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V28*/ meltfptr[24] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5832:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[8] = 0;
      /*^clear */
	     /*clear *//*_._IF___V28*/ meltfptr[24] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5834:/ locexp");

#if MELT_HAVE_DEBUG
      if (melt_need_debug (0))
	melt_dbgshortbacktrace (("normletrec_lambda ended"), (15));
#endif
      ;
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;


    MELT_LOCATION ("warmelt-normal.melt:5816:/ clear");
	   /*clear *//*_.CSYM__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.CLOCC__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.CONSTROUT__V20*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_.CLOVTUP__V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.SLOC__V22*/ meltfptr[21] = 0;
    /*^clear */
	   /*clear *//*_.OLDPROC__V23*/ meltfptr[22] = 0;
    /*^clear */
	   /*clear *//*_.NDATAROUT__V24*/ meltfptr[23] = 0;
    /*^clear */
	   /*clear *//*_#INSIDEFLAG__L7*/ meltfnum[0] = 0;

    MELT_LOCATION ("warmelt-normal.melt:5812:/ clear");
	   /*clear *//*_.NEWENV__V12*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V15*/ meltfptr[13] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */

    MELT_LOCATION ("warmelt-normal.melt:5808:/ clear");
	   /*clear *//*_.IFCPP___V10*/ meltfptr[8] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMLETREC_LAMBDA", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*noretval */ NULL);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_105_WARMELTmiNORMAL_NORMLETREC_LAMBDA */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *
					     meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un *
					     meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 20
    melt_ptr_t mcfr_varptr[20];
#define MELTFRAM_NBVARNUM 6
    long mcfr_varnum[6];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE_st *)
	meltfirstargp_;
      /* use arguments meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 20; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE nbval 20*/
  meltfram__.mcfr_nbvar = 20 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("PREPCONS_TUPLE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5840:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5841:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5841:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5841:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5841;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_tuple recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5841:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5841:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5842:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_TUPLE */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5842:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5842:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5842) ? (5842) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5842:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5843:/ quasiblock");


    MELT_LOCATION ("warmelt-normal.melt:5844:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_LOCATED */
						       meltfrout->tabval[2])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOC__V11*/ meltfptr[10] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LOC__V11*/ meltfptr[10] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5845:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_SOURCE_ARGUMENTED_OPERATOR */ meltfrout->tabval[3])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
   /*_.TUPARG__V12*/ meltfptr[11] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.TUPARG__V12*/ meltfptr[11] = NULL;;
      }
    ;
    /*^compute */
 /*_#NBTUPARG__L4*/ meltfnum[0] =
      (melt_multiple_length
       ((melt_ptr_t) ( /*_.TUPARG__V12*/ meltfptr[11])));;
    /*^compute */
 /*_.NTUP__V13*/ meltfptr[12] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[4])),
	( /*_#NBTUPARG__L4*/ meltfnum[0])));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5850:/ cond");
    /*cond */ if ( /*_.LOC__V11*/ meltfptr[10])	/*then */
      {
	/*^cond.then */
	/*_.OR___V14*/ meltfptr[13] = /*_.LOC__V11*/ meltfptr[10];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5850:/ cond.else");

	/*_.OR___V14*/ meltfptr[13] = /*_.SLOC__V5*/ meltfptr[4];;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5851:/ apply");
    /*apply */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^apply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V5*/ meltfptr[4];
      /*^apply.arg */
      argtab[2].meltbp_cstring = "discr_multiple";
      /*_.NORMAL_PREDEF__V15*/ meltfptr[14] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMAL_PREDEF */ meltfrout->tabval[6])),
		    (melt_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->
				   tabval[4])),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING ""),
		    argtab, "", (union meltparam_un *) 0);
    }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5848:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_TUPLE_BINDING */ meltfrout->tabval[5])), (5), "CLASS_NORMAL_CONSTRUCTED_TUPLE_BINDING");
  /*_.INST__V17*/ meltfptr[16] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @BINDER",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V17*/ meltfptr[16])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V17*/ meltfptr[16]), (0),
			  ( /*_.SYMB__V3*/ meltfptr[2]), "BINDER");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V17*/ meltfptr[16])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V17*/ meltfptr[16]), (1),
			  ( /*_.OR___V14*/ meltfptr[13]), "NCONSB_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V17*/ meltfptr[16])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V17*/ meltfptr[16]), (2),
			  ( /*_.NORMAL_PREDEF__V15*/ meltfptr[14]),
			  "NCONSB_DISCR");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NTUPB_COMP",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V17*/ meltfptr[16])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V17*/ meltfptr[16]), (4),
			  ( /*_.NTUP__V13*/ meltfptr[12]), "NTUPB_COMP");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V17*/ meltfptr[16],
				  "newly made instance");
    ;
    /*_.CONSTUP__V16*/ meltfptr[15] = /*_.INST__V17*/ meltfptr[16];;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5855:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L5*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5855:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L5*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L6*/ meltfnum[5] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5855:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L6*/ meltfnum[5];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5855;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_tuple gives constup";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSTUP__V16*/ meltfptr[15];
	      /*_.MELT_DEBUG_FUN__V19*/ meltfptr[18] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V18*/ meltfptr[17] =
	      /*_.MELT_DEBUG_FUN__V19*/ meltfptr[18];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5855:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L6*/ meltfnum[5] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V19*/ meltfptr[18] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V18*/ meltfptr[17] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5855:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L5*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V18*/ meltfptr[17] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5856:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.CONSTUP__V16*/ meltfptr[15];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5856:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V10*/ meltfptr[5] = /*_.RETURN___V20*/ meltfptr[18];;

    MELT_LOCATION ("warmelt-normal.melt:5843:/ clear");
	   /*clear *//*_.LOC__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.TUPARG__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_#NBTUPARG__L4*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.NTUP__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.OR___V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_.NORMAL_PREDEF__V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.CONSTUP__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V20*/ meltfptr[18] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5840:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V10*/ meltfptr[5];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5840:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.LET___V10*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("PREPCONS_TUPLE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_106_WARMELTmiNORMAL_PREPCONS_TUPLE */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE (meltclosure_ptr_t meltclosp_,
					       melt_ptr_t meltfirstargp_,
					       const melt_argdescr_cell_t
					       meltxargdescr_[],
					       union meltparam_un *
					       meltxargtab_,
					       const melt_argdescr_cell_t
					       meltxresdescr_[],
					       union meltparam_un *
					       meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 41
    melt_ptr_t mcfr_varptr[41];
#define MELTFRAM_NBVARNUM 20
    long mcfr_varnum[20];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE_st *)
	meltfirstargp_;
      /* use arguments meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 41; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE nbval 41*/
  meltfram__.mcfr_nbvar = 41 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMLETREC_TUPLE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5861:/ getarg");
 /*_.CEXPR__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.CBIND__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4])) != NULL);


  /*getarg#4 */
  /*^getarg */
  if (meltxargdescr_[3] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V6*/ meltfptr[5] =
    (meltxargtab_[3].meltbp_aptr) ? (*(meltxargtab_[3].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5])) != NULL);


  /*getarg#5 */
  /*^getarg */
  if (meltxargdescr_[4] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V7*/ meltfptr[6] =
    (meltxargtab_[4].meltbp_aptr) ? (*(meltxargtab_[4].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5862:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5862:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5862:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5862;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_tuple cexpr=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CEXPR__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " cbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V8*/ meltfptr[7] =
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5862:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V8*/ meltfptr[7] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5862:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5863:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_TUPLE */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5863:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[7] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5863:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cexpr"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5863) ? (5863) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[8] = /*_._IFELSE___V11*/ meltfptr[7];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5863:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[8] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5864:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:5864:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V13*/ meltfptr[12] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5864:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check env"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5864) ? (5864) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V12*/ meltfptr[7] = /*_._IFELSE___V13*/ meltfptr[12];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5864:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V12*/ meltfptr[7] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5865:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L5*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_TUPLE_BINDING */ meltfrout->tabval[3])));;
      MELT_LOCATION ("warmelt-normal.melt:5865:/ cond");
      /*cond */ if ( /*_#IS_A__L5*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V15*/ meltfptr[14] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5865:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cbind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5865) ? (5865) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V15*/ meltfptr[14] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V14*/ meltfptr[12] = /*_._IFELSE___V15*/ meltfptr[14];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5865:/ clear");
	     /*clear *//*_#IS_A__L5*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V15*/ meltfptr[14] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V14*/ meltfptr[12] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5866:/ quasiblock");


    MELT_LOCATION ("warmelt-normal.melt:5867:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CEXPR__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_LOCATED */
						       meltfrout->tabval[4])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOCA_LOCATION__V16*/ meltfptr[14] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LOCA_LOCATION__V16*/ meltfptr[14] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5867:/ cond");
    /*cond */ if ( /*_.LOCA_LOCATION__V16*/ meltfptr[14])	/*then */
      {
	/*^cond.then */
	/*_.LOC__V17*/ meltfptr[16] = /*_.LOCA_LOCATION__V16*/ meltfptr[14];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5867:/ cond.else");

	/*_.LOC__V17*/ meltfptr[16] = /*_.SLOC__V7*/ meltfptr[6];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5868:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
  /*_.SARGS__V18*/ meltfptr[17] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:5869:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CBIND__V4*/ meltfptr[3]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_TUPLE_BINDING */ meltfrout->tabval[3])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 4, "NTUPB_COMP");
   /*_.BCOMP__V19*/ meltfptr[18] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.BCOMP__V19*/ meltfptr[18] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5870:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CBIND__V4*/ meltfptr[3]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[5])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NCONSB_NLETREC");
   /*_.NLETREC__V20*/ meltfptr[19] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NLETREC__V20*/ meltfptr[19] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5871:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V20*/ meltfptr[19]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[6])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.NLETREC__V20*/ meltfptr[19]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NLETREC_FILL_BINDINGS");
   /*_.NRECBINDS__V21*/ meltfptr[20] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NRECBINDS__V21*/ meltfptr[20] = NULL;;
      }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5873:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L6*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5873:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L6*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L7*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5873:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[11];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L7*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5873;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_tuple sargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SARGS__V18*/ meltfptr[17];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " bcomp=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.BCOMP__V19*/ meltfptr[18];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " nletrec=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V20*/ meltfptr[19];
	      /*^apply.arg */
	      argtab[9].meltbp_cstring = " nrecbinds=";
	      /*^apply.arg */
	      argtab[10].meltbp_aptr =
		(melt_ptr_t *) & /*_.NRECBINDS__V21*/ meltfptr[20];
	      /*_.MELT_DEBUG_FUN__V23*/ meltfptr[22] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V22*/ meltfptr[21] =
	      /*_.MELT_DEBUG_FUN__V23*/ meltfptr[22];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5873:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L7*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V23*/ meltfptr[22] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V22*/ meltfptr[21] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5873:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L6*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V22*/ meltfptr[21] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5875:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L8*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NLETREC__V20*/ meltfptr[19]),
			     (melt_ptr_t) (( /*!CLASS_NREP_LETREC */
					    meltfrout->tabval[6])));;
      MELT_LOCATION ("warmelt-normal.melt:5875:/ cond");
      /*cond */ if ( /*_#IS_A__L8*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V25*/ meltfptr[21] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5875:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nletrec"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5875) ? (5875) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V25*/ meltfptr[21] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V24*/ meltfptr[22] = /*_._IFELSE___V25*/ meltfptr[21];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5875:/ clear");
	     /*clear *//*_#IS_A__L8*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V25*/ meltfptr[21] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V24*/ meltfptr[22] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5876:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST__L9*/ meltfnum[0] =
	(melt_magic_discr ((melt_ptr_t) ( /*_.NRECBINDS__V21*/ meltfptr[20]))
	 == MELTOBMAG_LIST);;
      MELT_LOCATION ("warmelt-normal.melt:5876:/ cond");
      /*cond */ if ( /*_#IS_LIST__L9*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V27*/ meltfptr[26] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5876:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nrecbinds"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5876) ? (5876) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V27*/ meltfptr[26] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V26*/ meltfptr[21] = /*_._IFELSE___V27*/ meltfptr[26];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5876:/ clear");
	     /*clear *//*_#IS_LIST__L9*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V27*/ meltfptr[26] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V26*/ meltfptr[21] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5877:/ quasiblock");


    /*^multiapply */
    /*multiapply 4args, 1x.res */
    {
      union meltparam_un argtab[3];

      union meltparam_un restab[1];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V6*/ meltfptr[5];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.LOC__V17*/ meltfptr[16];
      /*^multiapply.xres */
      restab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBINDS__V29*/ meltfptr[28];
      /*^multiapply.appl */
      /*_.NARGS__V28*/ meltfptr[26] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_TUPLE */ meltfrout->tabval[7])),
		    (melt_ptr_t) ( /*_.SARGS__V18*/ meltfptr[17]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		    argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5880:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L10*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5880:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L11*/ meltfnum[0] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5880:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L11*/ meltfnum[0];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5880;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_tuple nargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NARGS__V28*/ meltfptr[26];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nbinds=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBINDS__V29*/ meltfptr[28];
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[30] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V30*/ meltfptr[29] =
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[30];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5880:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[0] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V31*/ meltfptr[30] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V30*/ meltfptr[29] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5880:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V30*/ meltfptr[29] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5881:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#MULTIPLE_LENGTH__L12*/ meltfnum[0] =
	(melt_multiple_length
	 ((melt_ptr_t) ( /*_.BCOMP__V19*/ meltfptr[18])));;
      /*^compute */
   /*_#MULTIPLE_LENGTH__L13*/ meltfnum[1] =
	(melt_multiple_length
	 ((melt_ptr_t) ( /*_.NARGS__V28*/ meltfptr[26])));;
      /*^compute */
   /*_#eqeqI__L14*/ meltfnum[13] =
	(( /*_#MULTIPLE_LENGTH__L12*/ meltfnum[0]) ==
	 ( /*_#MULTIPLE_LENGTH__L13*/ meltfnum[1]));;
      MELT_LOCATION ("warmelt-normal.melt:5881:/ cond");
      /*cond */ if ( /*_#eqeqI__L14*/ meltfnum[13])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V33*/ meltfptr[29] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5881:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nargs & bcomp samelength"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5881) ? (5881) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V33*/ meltfptr[29] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V32*/ meltfptr[30] = /*_._IFELSE___V33*/ meltfptr[29];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5881:/ clear");
	     /*clear *//*_#MULTIPLE_LENGTH__L12*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_#MULTIPLE_LENGTH__L13*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_#eqeqI__L14*/ meltfnum[13] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V33*/ meltfptr[29] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V32*/ meltfptr[30] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5882:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L15*/ meltfnum[0] =
	(( /*_.NBINDS__V29*/ meltfptr[28]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBINDS__V29*/ meltfptr[28])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:5882:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L15*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V35*/ meltfptr[34] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5882:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nbinds"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5882) ? (5882) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V35*/ meltfptr[34] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V34*/ meltfptr[29] = /*_._IFELSE___V35*/ meltfptr[34];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5882:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L15*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V35*/ meltfptr[34] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V34*/ meltfptr[29] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NARGS__V28*/ meltfptr[26]);
      for ( /*_#NIX__L16*/ meltfnum[1] = 0;
	   ( /*_#NIX__L16*/ meltfnum[1] >= 0)
	   && ( /*_#NIX__L16*/ meltfnum[1] < meltcit1__EACHTUP_ln);
	/*_#NIX__L16*/ meltfnum[1]++)
	{
	  /*_.CURNARG__V36*/ meltfptr[34] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.NARGS__V28*/ meltfptr[26]),
			       /*_#NIX__L16*/ meltfnum[1]);




	  {
	    MELT_LOCATION ("warmelt-normal.melt:5886:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.BCOMP__V19*/ meltfptr[18]),
				     ( /*_#NIX__L16*/ meltfnum[1]),
				     (melt_ptr_t) ( /*_.CURNARG__V36*/
						   meltfptr[34]));
	  }
	  ;
	  if ( /*_#NIX__L16*/ meltfnum[1] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5883:/ clear");
	    /*clear *//*_.CURNARG__V36*/ meltfptr[34] = 0;
      /*^clear */
	    /*clear *//*_#NIX__L16*/ meltfnum[1] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5887:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L17*/ meltfnum[13] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5887:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L17*/ meltfnum[13])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L18*/ meltfnum[0] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5887:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L18*/ meltfnum[0];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5887;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_tuple updated bcomp";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.BCOMP__V19*/ meltfptr[18];
	      /*_.MELT_DEBUG_FUN__V38*/ meltfptr[37] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V37*/ meltfptr[36] =
	      /*_.MELT_DEBUG_FUN__V38*/ meltfptr[37];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5887:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L18*/ meltfnum[0] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V38*/ meltfptr[37] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V37*/ meltfptr[36] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5887:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L17*/ meltfnum[13] = 0;
      /*^clear */
	     /*clear *//*_._IF___V37*/ meltfptr[36] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5888:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBINDS__V29*/ meltfptr[28];
      /*_.LIST_APPEND2LIST__V39*/ meltfptr[37] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_APPEND2LIST */ meltfrout->tabval[8])),
		    (melt_ptr_t) ( /*_.NRECBINDS__V21*/ meltfptr[20]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5889:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L19*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5889:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L19*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L20*/ meltfnum[13] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5889:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L20*/ meltfnum[13];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5889;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_tuple updated nrecbinds";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NRECBINDS__V21*/ meltfptr[20];
	      /*_.MELT_DEBUG_FUN__V41*/ meltfptr[40] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V40*/ meltfptr[36] =
	      /*_.MELT_DEBUG_FUN__V41*/ meltfptr[40];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5889:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L20*/ meltfnum[13] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V41*/ meltfptr[40] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V40*/ meltfptr[36] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5889:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L19*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V40*/ meltfptr[36] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5890:/ locexp");

#if MELT_HAVE_DEBUG
      if (melt_need_debug (0))
	melt_dbgshortbacktrace (("normletrec_tuple ended"), (15));
#endif
      ;
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;

    MELT_LOCATION ("warmelt-normal.melt:5877:/ clear");
	   /*clear *//*_.IFCPP___V32*/ meltfptr[30] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V34*/ meltfptr[29] = 0;
    /*^clear */
	   /*clear *//*_.LIST_APPEND2LIST__V39*/ meltfptr[37] = 0;

    /*^clear */
	   /*clear *//*_.NBINDS__V29*/ meltfptr[28] = 0;

    MELT_LOCATION ("warmelt-normal.melt:5866:/ clear");
	   /*clear *//*_.LOCA_LOCATION__V16*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.LOC__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.SARGS__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.BCOMP__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.NLETREC__V20*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_.NRECBINDS__V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V24*/ meltfptr[22] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V26*/ meltfptr[21] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */

    MELT_LOCATION ("warmelt-normal.melt:5861:/ clear");
	   /*clear *//*_.IFCPP___V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V12*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V14*/ meltfptr[12] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMLETREC_TUPLE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*noretval */ NULL);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_107_WARMELTmiNORMAL_NORMLETREC_TUPLE */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un * meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 32
    melt_ptr_t mcfr_varptr[32];
#define MELTFRAM_NBVARNUM 10
    long mcfr_varnum[10];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST_st *)
	meltfirstargp_;
      /* use arguments meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 32; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST nbval 32*/
  meltfram__.mcfr_nbvar = 32 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("PREPCONS_LIST", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5896:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5897:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5897:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5897:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5897;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_list recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5897:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5897:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5898:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_LIST */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5898:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5898:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5898) ? (5898) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5898:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5899:/ quasiblock");


 /*_.RESLIST__V11*/ meltfptr[10] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[2]))));;
    MELT_LOCATION ("warmelt-normal.melt:5901:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_LOCATED */
						       meltfrout->tabval[3])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOC__V12*/ meltfptr[11] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LOC__V12*/ meltfptr[11] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5902:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_SOURCE_ARGUMENTED_OPERATOR */ meltfrout->tabval[4])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
   /*_.LISTARG__V13*/ meltfptr[12] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LISTARG__V13*/ meltfptr[12] = NULL;;
      }
    ;
    /*^compute */
 /*_#NBLISTARG__L4*/ meltfnum[0] =
      (melt_multiple_length
       ((melt_ptr_t) ( /*_.LISTARG__V13*/ meltfptr[12])));;
    /*^compute */
 /*_.PAIRSB__V14*/ meltfptr[13] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[5])),
	( /*_#NBLISTARG__L4*/ meltfnum[0])));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5907:/ cond");
    /*cond */ if ( /*_.LOC__V12*/ meltfptr[11])	/*then */
      {
	/*^cond.then */
	/*_.OR___V15*/ meltfptr[14] = /*_.LOC__V12*/ meltfptr[11];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5907:/ cond.else");

	/*_.OR___V15*/ meltfptr[14] = /*_.SLOC__V5*/ meltfptr[4];;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5908:/ apply");
    /*apply */
    {
      union meltparam_un argtab[3];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
      /*^apply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V5*/ meltfptr[4];
      /*^apply.arg */
      argtab[2].meltbp_cstring = "discr_list";
      /*_.NORMAL_PREDEF__V16*/ meltfptr[15] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMAL_PREDEF */ meltfrout->tabval[7])),
		    (melt_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[2])),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING ""),
		    argtab, "", (union meltparam_un *) 0);
    }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5905:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LIST_BINDING */ meltfrout->tabval[6])), (7), "CLASS_NORMAL_CONSTRUCTED_LIST_BINDING");
  /*_.INST__V18*/ meltfptr[17] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @BINDER",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (0),
			  ( /*_.SYMB__V3*/ meltfptr[2]), "BINDER");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (1),
			  ( /*_.OR___V15*/ meltfptr[14]), "NCONSB_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (2),
			  ( /*_.NORMAL_PREDEF__V16*/ meltfptr[15]),
			  "NCONSB_DISCR");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLISTB_FIRST",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (4),
			  (( /*nil */ NULL)), "NLISTB_FIRST");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLISTB_LAST",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (5),
			  (( /*nil */ NULL)), "NLISTB_LAST");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLISTB_PAIRSB",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V18*/ meltfptr[17])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V18*/ meltfptr[17]), (6),
			  ( /*_.PAIRSB__V14*/ meltfptr[13]), "NLISTB_PAIRSB");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V18*/ meltfptr[17],
				  "newly made instance");
    ;
    /*_.CONSLIST__V17*/ meltfptr[16] = /*_.INST__V18*/ meltfptr[17];;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.LISTARG__V13*/ meltfptr[12]);
      for ( /*_#CURIX__L5*/ meltfnum[1] = 0;
	   ( /*_#CURIX__L5*/ meltfnum[1] >= 0)
	   && ( /*_#CURIX__L5*/ meltfnum[1] < meltcit1__EACHTUP_ln);
	/*_#CURIX__L5*/ meltfnum[1]++)
	{
	  /*_.CURARG__V19*/ meltfptr[18] =
	    melt_multiple_nth ((melt_ptr_t)
			       ( /*_.LISTARG__V13*/ meltfptr[12]),
			       /*_#CURIX__L5*/ meltfnum[1]);




#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:5917:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L6*/ meltfnum[5] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:5917:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L6*/ meltfnum[5])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L7*/ meltfnum[6] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5917:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L7*/ meltfnum[6];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 5917;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "prepcons_list curarg";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURARG__V19*/ meltfptr[18];
		    /*_.MELT_DEBUG_FUN__V21*/ meltfptr[20] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V20*/ meltfptr[19] =
		    /*_.MELT_DEBUG_FUN__V21*/ meltfptr[20];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5917:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L7*/ meltfnum[6] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V21*/ meltfptr[20] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V20*/ meltfptr[19] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:5917:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L6*/ meltfnum[5] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V20*/ meltfptr[19] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5918:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
  /*_#IS_A__L8*/ meltfnum[6] =
	    melt_is_instance_of ((melt_ptr_t)
				 ( /*_.CURARG__V19*/ meltfptr[18]),
				 (melt_ptr_t) (( /*!CLASS_NAMED */ meltfrout->
						tabval[9])));;
	  MELT_LOCATION ("warmelt-normal.melt:5918:/ cond");
	  /*cond */ if ( /*_#IS_A__L8*/ meltfnum[6])	/*then */
	    {
	      /*^cond.then */
	      /*_._IFELSE___V22*/ meltfptr[20] =
		/*_.CURARG__V19*/ meltfptr[18];;
	    }
	  else
	    {
	      MELT_LOCATION ("warmelt-normal.melt:5918:/ cond.else");

	      /*_._IFELSE___V22*/ meltfptr[20] =
		( /*!konst_10__PAIROFLIST */ meltfrout->tabval[10]);;
	    }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5918:/ apply");
	  /*apply */
	  {
	    /*_.PAIRSYMB__V23*/ meltfptr[19] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!CLONE_SYMBOL */ meltfrout->tabval[8])),
			  (melt_ptr_t) ( /*_._IFELSE___V22*/ meltfptr[20]),
			  (""), (union meltparam_un *) 0, "",
			  (union meltparam_un *) 0);
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5921:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURARG__V19*/
					       meltfptr[18]),
					      (melt_ptr_t) (( /*!CLASS_LOCATED */ meltfrout->tabval[3])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURARG__V19*/ meltfptr[18]) /*=obj*/ ;
		melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
    /*_.LOCA_LOCATION__V24*/ meltfptr[23] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.LOCA_LOCATION__V24*/ meltfptr[23] = NULL;;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5921:/ cond");
	  /*cond */ if ( /*_.LOCA_LOCATION__V24*/ meltfptr[23])	/*then */
	    {
	      /*^cond.then */
	      /*_.OR___V25*/ meltfptr[24] =
		/*_.LOCA_LOCATION__V24*/ meltfptr[23];;
	    }
	  else
	    {
	      MELT_LOCATION ("warmelt-normal.melt:5921:/ cond.else");

	      /*^block */
	      /*anyblock */
	      {

		/*^cond */
		/*cond */ if ( /*_.LOC__V12*/ meltfptr[11])	/*then */
		  {
		    /*^cond.then */
		    /*_.OR___V26*/ meltfptr[25] =
		      /*_.LOC__V12*/ meltfptr[11];;
		  }
		else
		  {
		    MELT_LOCATION ("warmelt-normal.melt:5921:/ cond.else");

		    /*_.OR___V26*/ meltfptr[25] = /*_.SLOC__V5*/ meltfptr[4];;
		  }
		;
		/*^compute */
		/*_.OR___V25*/ meltfptr[24] = /*_.OR___V26*/ meltfptr[25];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5921:/ clear");
	      /*clear *//*_.OR___V26*/ meltfptr[25] = 0;
	      }
	      ;
	    }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5922:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V5*/ meltfptr[4];
	    /*^apply.arg */
	    argtab[2].meltbp_cstring = "discr_pair";
	    /*_.NORMAL_PREDEF__V27*/ meltfptr[25] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_PREDEF */ meltfrout->tabval[7])),
			  (melt_ptr_t) (( /*!DISCR_PAIR */ meltfrout->
					 tabval[12])),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			   ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5919:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[11])), (6), "CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING");
   /*_.INST__V29*/ meltfptr[28] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @BINDER",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V29*/ meltfptr[28]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V29*/ meltfptr[28]), (0),
				( /*_.PAIRSYMB__V23*/ meltfptr[19]),
				"BINDER");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NCONSB_LOC",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V29*/ meltfptr[28]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V29*/ meltfptr[28]), (1),
				( /*_.OR___V25*/ meltfptr[24]), "NCONSB_LOC");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V29*/ meltfptr[28]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V29*/ meltfptr[28]), (2),
				( /*_.NORMAL_PREDEF__V27*/ meltfptr[25]),
				"NCONSB_DISCR");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NPAIRB_HEAD",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V29*/ meltfptr[28]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V29*/ meltfptr[28]), (4),
				(( /*nil */ NULL)), "NPAIRB_HEAD");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NPAIRB_TAIL",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V29*/ meltfptr[28]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V29*/ meltfptr[28]), (5),
				(( /*nil */ NULL)), "NPAIRB_TAIL");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V29*/ meltfptr[28],
					"newly made instance");
	  ;
	  /*_.CONSPAIR__V28*/ meltfptr[27] = /*_.INST__V29*/ meltfptr[28];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:5927:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.PAIRSB__V14*/ meltfptr[13]),
				     ( /*_#CURIX__L5*/ meltfnum[1]),
				     (melt_ptr_t) ( /*_.CONSPAIR__V28*/
						   meltfptr[27]));
	  }
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:5928:/ locexp");
	    meltgc_append_list ((melt_ptr_t)
				( /*_.RESLIST__V11*/ meltfptr[10]),
				(melt_ptr_t) ( /*_.CONSPAIR__V28*/
					      meltfptr[27]));
	  }
	  ;

	  MELT_LOCATION ("warmelt-normal.melt:5918:/ clear");
	    /*clear *//*_#IS_A__L8*/ meltfnum[6] = 0;
	  /*^clear */
	    /*clear *//*_._IFELSE___V22*/ meltfptr[20] = 0;
	  /*^clear */
	    /*clear *//*_.PAIRSYMB__V23*/ meltfptr[19] = 0;
	  /*^clear */
	    /*clear *//*_.LOCA_LOCATION__V24*/ meltfptr[23] = 0;
	  /*^clear */
	    /*clear *//*_.OR___V25*/ meltfptr[24] = 0;
	  /*^clear */
	    /*clear *//*_.NORMAL_PREDEF__V27*/ meltfptr[25] = 0;
	  /*^clear */
	    /*clear *//*_.CONSPAIR__V28*/ meltfptr[27] = 0;
	  if ( /*_#CURIX__L5*/ meltfnum[1] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5914:/ clear");
	    /*clear *//*_.CURARG__V19*/ meltfptr[18] = 0;
      /*^clear */
	    /*clear *//*_#CURIX__L5*/ meltfnum[1] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:5930:/ locexp");
      meltgc_append_list ((melt_ptr_t) ( /*_.RESLIST__V11*/ meltfptr[10]),
			  (melt_ptr_t) ( /*_.CONSLIST__V17*/ meltfptr[16]));
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5931:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L9*/ meltfnum[5] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5931:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L9*/ meltfnum[5])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L10*/ meltfnum[6] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5931:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L10*/ meltfnum[6];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5931;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_list gives reslist";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RESLIST__V11*/ meltfptr[10];
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[19] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V30*/ meltfptr[20] =
	      /*_.MELT_DEBUG_FUN__V31*/ meltfptr[19];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5931:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L10*/ meltfnum[6] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V31*/ meltfptr[19] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V30*/ meltfptr[20] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5931:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L9*/ meltfnum[5] = 0;
      /*^clear */
	     /*clear *//*_._IF___V30*/ meltfptr[20] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5932:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.RESLIST__V11*/ meltfptr[10];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5932:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V10*/ meltfptr[5] = /*_.RETURN___V32*/ meltfptr[23];;

    MELT_LOCATION ("warmelt-normal.melt:5899:/ clear");
	   /*clear *//*_.RESLIST__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.LOC__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_.LISTARG__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_#NBLISTARG__L4*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.PAIRSB__V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_.OR___V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.NORMAL_PREDEF__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.CONSLIST__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V32*/ meltfptr[23] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5896:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V10*/ meltfptr[5];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5896:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.LET___V10*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("PREPCONS_LIST", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_108_WARMELTmiNORMAL_PREPCONS_LIST */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM (meltclosure_ptr_t
						   meltclosp_,
						   melt_ptr_t meltfirstargp_,
						   const melt_argdescr_cell_t
						   meltxargdescr_[],
						   union meltparam_un *
						   meltxargtab_,
						   const melt_argdescr_cell_t
						   meltxresdescr_[],
						   union meltparam_un *
						   meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 21
    melt_ptr_t mcfr_varptr[21];
#define MELTFRAM_NBVARNUM 10
    long mcfr_varnum[10];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM_st
	 *) meltfirstargp_;
      /* use arguments meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 21; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM nbval 21*/
  meltfram__.mcfr_nbvar = 21 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NRECLIST_FIND_LOCSYM", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5939:/ getarg");
 /*_.SYMB__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NLETREC__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NLETREC__V3*/ meltfptr[2])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5940:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5940:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5940:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5940;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "nreclist_find_locsym symb";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V5*/ meltfptr[4] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V4*/ meltfptr[3] =
	      /*_.MELT_DEBUG_FUN__V5*/ meltfptr[4];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5940:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V5*/ meltfptr[4] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V4*/ meltfptr[3] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5940:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V4*/ meltfptr[3] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5941:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.SYMB__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SYMBOL */ meltfrout->
					    tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5941:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V7*/ meltfptr[3] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5941:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check symb"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5941) ? (5941) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V7*/ meltfptr[3] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V6*/ meltfptr[4] = /*_._IFELSE___V7*/ meltfptr[3];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5941:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V7*/ meltfptr[3] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V6*/ meltfptr[4] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5942:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NLETREC__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_NREP_LETREC */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:5942:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5942:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nletrec"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5942) ? (5942) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[3] = /*_._IFELSE___V9*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5942:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[3] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5943:/ quasiblock");


    /*^cond */
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V3*/ meltfptr[2]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[2])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.NLETREC__V3*/ meltfptr[2]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 5, "NLETREC_LOCSYMS");
   /*_.NLOCSYMS__V11*/ meltfptr[10] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NLOCSYMS__V11*/ meltfptr[10] = NULL;;
      }
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NLOCSYMS__V11*/ meltfptr[10]);
      for ( /*_#LOCSYMIX__L5*/ meltfnum[1] = 0;
	   ( /*_#LOCSYMIX__L5*/ meltfnum[1] >= 0)
	   && ( /*_#LOCSYMIX__L5*/ meltfnum[1] < meltcit1__EACHTUP_ln);
	/*_#LOCSYMIX__L5*/ meltfnum[1]++)
	{
	  /*_.CURLOCSYM__V12*/ meltfptr[11] =
	    melt_multiple_nth ((melt_ptr_t)
			       ( /*_.NLOCSYMS__V11*/ meltfptr[10]),
			       /*_#LOCSYMIX__L5*/ meltfnum[1]);




	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5948:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURLOCSYM__V12*/
					       meltfptr[11]),
					      (melt_ptr_t) (( /*!CLASS_NREP_SYMOCC */ meltfrout->tabval[3])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURLOCSYM__V12*/ meltfptr[11]) /*=obj*/ ;
		melt_object_get_field (slot, obj, 1, "NOCC_SYMB");
    /*_.NOCC_SYMB__V13*/ meltfptr[12] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.NOCC_SYMB__V13*/ meltfptr[12] = NULL;;
	    }
	  ;
	  /*^compute */
  /*_#eqeq__L6*/ meltfnum[0] =
	    (( /*_.NOCC_SYMB__V13*/ meltfptr[12]) ==
	     ( /*_.SYMB__V2*/ meltfptr[1]));;
	  MELT_LOCATION ("warmelt-normal.melt:5948:/ cond");
	  /*cond */ if ( /*_#eqeq__L6*/ meltfnum[0])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {


#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:5950:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#MELT_NEED_DBG__L7*/ meltfnum[6] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:5950:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L7*/ meltfnum[6])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	/*_#MELT_CALLCOUNT__L8*/ meltfnum[7] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:5950:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[5];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L8*/ meltfnum[7];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 5950;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "nreclist_find_locsym found curlocsym";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CURLOCSYM__V12*/
			    meltfptr[11];
			  /*_.MELT_DEBUG_FUN__V16*/ meltfptr[15] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[0])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V15*/ meltfptr[14] =
			  /*_.MELT_DEBUG_FUN__V16*/ meltfptr[15];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:5950:/ clear");
		  /*clear *//*_#MELT_CALLCOUNT__L8*/ meltfnum[7] = 0;
			/*^clear */
		  /*clear *//*_.MELT_DEBUG_FUN__V16*/ meltfptr[15] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

       /*_._IF___V15*/ meltfptr[14] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:5950:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		/*clear *//*_#MELT_NEED_DBG__L7*/ meltfnum[6] = 0;
		  /*^clear */
		/*clear *//*_._IF___V15*/ meltfptr[14] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:5951:/ quasiblock");


		/*_._RETVAL___V1*/ meltfptr[0] =
		  /*_.CURLOCSYM__V12*/ meltfptr[11];;

		{
		  MELT_LOCATION ("warmelt-normal.melt:5951:/ locexp");
		  /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
		  if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
		    melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
		  ;
		}
		;
		/*^finalreturn */
		;
		/*finalret */ goto meltlabend_rout;
		MELT_LOCATION ("warmelt-normal.melt:5949:/ quasiblock");


		/*_.PROGN___V18*/ meltfptr[14] =
		  /*_.RETURN___V17*/ meltfptr[15];;
		/*^compute */
		/*_._IF___V14*/ meltfptr[13] =
		  /*_.PROGN___V18*/ meltfptr[14];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5948:/ clear");
	      /*clear *//*_.RETURN___V17*/ meltfptr[15] = 0;
		/*^clear */
	      /*clear *//*_.PROGN___V18*/ meltfptr[14] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_._IF___V14*/ meltfptr[13] = NULL;;
	    }
	  ;
	  if ( /*_#LOCSYMIX__L5*/ meltfnum[1] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5945:/ clear");
	    /*clear *//*_.CURLOCSYM__V12*/ meltfptr[11] = 0;
      /*^clear */
	    /*clear *//*_#LOCSYMIX__L5*/ meltfnum[1] = 0;
      /*^clear */
	    /*clear *//*_.NOCC_SYMB__V13*/ meltfptr[12] = 0;
      /*^clear */
	    /*clear *//*_#eqeq__L6*/ meltfnum[0] = 0;
      /*^clear */
	    /*clear *//*_._IF___V14*/ meltfptr[13] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5952:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L9*/ meltfnum[7] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5952:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L9*/ meltfnum[7])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L10*/ meltfnum[6] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5952:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[4];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L10*/ meltfnum[6];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5952;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "nreclist_find_locsym not found";
	      /*_.MELT_DEBUG_FUN__V20*/ meltfptr[14] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING ""), argtab,
			    "", (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V19*/ meltfptr[15] =
	      /*_.MELT_DEBUG_FUN__V20*/ meltfptr[14];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5952:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L10*/ meltfnum[6] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V20*/ meltfptr[14] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V19*/ meltfptr[15] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5952:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L9*/ meltfnum[7] = 0;
      /*^clear */
	     /*clear *//*_._IF___V19*/ meltfptr[15] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5953:/ quasiblock");


 /*_._RETVAL___V1*/ meltfptr[0] = NULL;;

    {
      MELT_LOCATION ("warmelt-normal.melt:5953:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V10*/ meltfptr[8] = /*_.RETURN___V21*/ meltfptr[14];;

    MELT_LOCATION ("warmelt-normal.melt:5943:/ clear");
	   /*clear *//*_.NLOCSYMS__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V21*/ meltfptr[14] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5939:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V10*/ meltfptr[8];;

    {
      MELT_LOCATION ("warmelt-normal.melt:5939:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V6*/ meltfptr[4] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[3] = 0;
    /*^clear */
	   /*clear *//*_.LET___V10*/ meltfptr[8] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NRECLIST_FIND_LOCSYM", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_109_WARMELTmiNORMAL_NRECLIST_FIND_LOCSYM */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST (meltclosure_ptr_t meltclosp_,
					      melt_ptr_t meltfirstargp_,
					      const melt_argdescr_cell_t
					      meltxargdescr_[],
					      union meltparam_un *
					      meltxargtab_,
					      const melt_argdescr_cell_t
					      meltxresdescr_[],
					      union meltparam_un *
					      meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 76
    melt_ptr_t mcfr_varptr[76];
#define MELTFRAM_NBVARNUM 43
    long mcfr_varnum[43];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST_st *)
	meltfirstargp_;
      /* use arguments meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 76; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST nbval 76*/
  meltfram__.mcfr_nbvar = 76 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMLETREC_LIST", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:5957:/ getarg");
 /*_.CEXPR__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.CBIND__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4])) != NULL);


  /*getarg#4 */
  /*^getarg */
  if (meltxargdescr_[3] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V6*/ meltfptr[5] =
    (meltxargtab_[3].meltbp_aptr) ? (*(meltxargtab_[3].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5])) != NULL);


  /*getarg#5 */
  /*^getarg */
  if (meltxargdescr_[4] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V7*/ meltfptr[6] =
    (meltxargtab_[4].meltbp_aptr) ? (*(meltxargtab_[4].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5958:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5958:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5958:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5958;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list cexpr=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CEXPR__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " cbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V8*/ meltfptr[7] =
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5958:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V8*/ meltfptr[7] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5958:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5959:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_LIST */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:5959:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[7] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5959:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cexpr"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5959) ? (5959) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[8] = /*_._IFELSE___V11*/ meltfptr[7];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5959:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[8] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5962:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_MULTIPLE__L4*/ meltfnum[0] =
	(melt_magic_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) ==
	 MELTOBMAG_MULTIPLE);;
      MELT_LOCATION ("warmelt-normal.melt:5962:/ cond");
      /*cond */ if ( /*_#IS_MULTIPLE__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V13*/ meltfptr[12] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5962:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cbind is tuple"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5962) ? (5962) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V12*/ meltfptr[7] = /*_._IFELSE___V13*/ meltfptr[12];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5962:/ clear");
	     /*clear *//*_#IS_MULTIPLE__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V12*/ meltfptr[7] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:5963:/ quasiblock");


    /*^cond */
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CEXPR__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_LOCATED */
						       meltfrout->tabval[2])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOCA_LOCATION__V14*/ meltfptr[12] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LOCA_LOCATION__V14*/ meltfptr[12] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5963:/ cond");
    /*cond */ if ( /*_.LOCA_LOCATION__V14*/ meltfptr[12])	/*then */
      {
	/*^cond.then */
	/*_.LOC__V15*/ meltfptr[14] = /*_.LOCA_LOCATION__V14*/ meltfptr[12];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:5963:/ cond.else");

	/*_.LOC__V15*/ meltfptr[14] = /*_.SLOC__V7*/ meltfptr[6];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5964:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CEXPR__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_SOURCE_ARGUMENTED_OPERATOR */ meltfrout->tabval[3])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 2, "SARGOP_ARGS");
   /*_.SARGS__V16*/ meltfptr[15] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.SARGS__V16*/ meltfptr[15] = NULL;;
      }
    ;
    /*^compute */
 /*_.CONSLIBIND__V17*/ meltfptr[16] =
      (melt_multiple_nth
       ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]), (-1)));;
    MELT_LOCATION ("warmelt-normal.melt:5966:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CONSLIBIND__V17*/ meltfptr[16]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[4])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CONSLIBIND__V17*/ meltfptr[16]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NCONSB_NLETREC");
   /*_.NLETREC__V18*/ meltfptr[17] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NLETREC__V18*/ meltfptr[17] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5967:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V18*/ meltfptr[17]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[5])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.NLETREC__V18*/ meltfptr[17]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NLETREC_FILL_BINDINGS");
   /*_.NRECBINDS__V19*/ meltfptr[18] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NRECBINDS__V19*/ meltfptr[18] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:5968:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V18*/ meltfptr[17]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[5])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.NLETREC__V18*/ meltfptr[17]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 5, "NLETREC_LOCSYMS");
   /*_.NLOCSYMS__V20*/ meltfptr[19] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NLOCSYMS__V20*/ meltfptr[19] = NULL;;
      }
    ;
    /*^compute */
 /*_.FIRSTPAIRBIND__V21*/ meltfptr[20] =
      (melt_multiple_nth ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]), (0)));;
    /*^compute */
 /*_.LASTPAIRBIND__V22*/ meltfptr[21] =
      (melt_multiple_nth
       ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]), (-2)));;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5972:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L5*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5972:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L5*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L6*/ meltfnum[0] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5972:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L6*/ meltfnum[0];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5972;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list conslibind";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSLIBIND__V17*/ meltfptr[16];
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V23*/ meltfptr[22] =
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5972:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L6*/ meltfnum[0] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V23*/ meltfptr[22] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5972:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L5*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V23*/ meltfptr[22] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5973:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L7*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t)
			     ( /*_.CONSLIBIND__V17*/ meltfptr[16]),
			     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LIST_BINDING */ meltfrout->tabval[6])));;
      MELT_LOCATION ("warmelt-normal.melt:5973:/ cond");
      /*cond */ if ( /*_#IS_A__L7*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V26*/ meltfptr[22] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5973:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check conslibind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5973) ? (5973) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V26*/ meltfptr[22] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V25*/ meltfptr[23] = /*_._IFELSE___V26*/ meltfptr[22];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5973:/ clear");
	     /*clear *//*_#IS_A__L7*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V26*/ meltfptr[22] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V25*/ meltfptr[23] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5975:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L8*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5975:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L8*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L9*/ meltfnum[0] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5975:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L9*/ meltfnum[0];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5975;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list sargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SARGS__V16*/ meltfptr[15];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nletrec=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V18*/ meltfptr[17];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " nrecbinds=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.NRECBINDS__V19*/ meltfptr[18];
	      /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V27*/ meltfptr[22] =
	      /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5975:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L9*/ meltfnum[0] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V27*/ meltfptr[22] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5975:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L8*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V27*/ meltfptr[22] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5977:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L10*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NLETREC__V18*/ meltfptr[17]),
			     (melt_ptr_t) (( /*!CLASS_NREP_LETREC */
					    meltfrout->tabval[5])));;
      MELT_LOCATION ("warmelt-normal.melt:5977:/ cond");
      /*cond */ if ( /*_#IS_A__L10*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V30*/ meltfptr[22] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5977:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nletrec"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5977) ? (5977) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V30*/ meltfptr[22] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V29*/ meltfptr[27] = /*_._IFELSE___V30*/ meltfptr[22];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5977:/ clear");
	     /*clear *//*_#IS_A__L10*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V30*/ meltfptr[22] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V29*/ meltfptr[27] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5978:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST__L11*/ meltfnum[1] =
	(melt_magic_discr ((melt_ptr_t) ( /*_.NRECBINDS__V19*/ meltfptr[18]))
	 == MELTOBMAG_LIST);;
      MELT_LOCATION ("warmelt-normal.melt:5978:/ cond");
      /*cond */ if ( /*_#IS_LIST__L11*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V32*/ meltfptr[31] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5978:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nrecbinds"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5978) ? (5978) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V32*/ meltfptr[31] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V31*/ meltfptr[22] = /*_._IFELSE___V32*/ meltfptr[31];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5978:/ clear");
	     /*clear *//*_#IS_LIST__L11*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V32*/ meltfptr[31] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V31*/ meltfptr[22] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:5979:/ quasiblock");


    /*^multiapply */
    /*multiapply 4args, 1x.res */
    {
      union meltparam_un argtab[3];

      union meltparam_un restab[1];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V6*/ meltfptr[5];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.LOC__V15*/ meltfptr[14];
      /*^multiapply.xres */
      restab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBINDS__V34*/ meltfptr[33];
      /*^multiapply.appl */
      /*_.NARGS__V33*/ meltfptr[31] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_TUPLE */ meltfrout->tabval[7])),
		    (melt_ptr_t) ( /*_.SARGS__V16*/ meltfptr[15]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		    argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5982:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L12*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:5982:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L12*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L13*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:5982:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L13*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 5982;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list nargs=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NARGS__V33*/ meltfptr[31];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nbinds=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBINDS__V34*/ meltfptr[33];
	      /*_.MELT_DEBUG_FUN__V36*/ meltfptr[35] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V35*/ meltfptr[34] =
	      /*_.MELT_DEBUG_FUN__V36*/ meltfptr[35];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5982:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L13*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V36*/ meltfptr[35] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V35*/ meltfptr[34] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:5982:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L12*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V35*/ meltfptr[34] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5983:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L14*/ meltfnum[1] =
	(( /*_.NBINDS__V34*/ meltfptr[33]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBINDS__V34*/ meltfptr[33])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:5983:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L14*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V38*/ meltfptr[34] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5983:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nbinds"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (5983) ? (5983) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V38*/ meltfptr[34] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V37*/ meltfptr[35] = /*_._IFELSE___V38*/ meltfptr[34];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5983:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L14*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V38*/ meltfptr[34] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V37*/ meltfptr[35] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:5984:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#MULTIPLE_LENGTH__L15*/ meltfnum[0] =
	(melt_multiple_length
	 ((melt_ptr_t) ( /*_.NARGS__V33*/ meltfptr[31])));;
      /*^compute */
   /*_#MULTIPLE_LENGTH__L16*/ meltfnum[1] =
	(melt_multiple_length ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])));;
      /*^compute */
   /*_#miI__L17*/ meltfnum[16] =
	(( /*_#MULTIPLE_LENGTH__L16*/ meltfnum[1]) - (1));;
      /*^compute */
   /*_#eqeqI__L18*/ meltfnum[17] =
	(( /*_#MULTIPLE_LENGTH__L15*/ meltfnum[0]) ==
	 ( /*_#miI__L17*/ meltfnum[16]));;
      MELT_LOCATION ("warmelt-normal.melt:5984:/ cond");
      /*cond */ if ( /*_#eqeqI__L18*/ meltfnum[17])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V40*/ meltfptr[39] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:5984:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nargs\'s length vs cbind\'s length"), ("warmelt-normal.melt") ? ("warmelt-normal.melt") : __FILE__, (5984) ? (5984) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V40*/ meltfptr[39] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V39*/ meltfptr[34] = /*_._IFELSE___V40*/ meltfptr[39];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:5984:/ clear");
	     /*clear *//*_#MULTIPLE_LENGTH__L15*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_#MULTIPLE_LENGTH__L16*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_#miI__L17*/ meltfnum[16] = 0;
      /*^clear */
	     /*clear *//*_#eqeqI__L18*/ meltfnum[17] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V40*/ meltfptr[39] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V39*/ meltfptr[34] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NARGS__V33*/ meltfptr[31]);
      for ( /*_#NIX__L19*/ meltfnum[0] = 0;
	   ( /*_#NIX__L19*/ meltfnum[0] >= 0)
	   && ( /*_#NIX__L19*/ meltfnum[0] < meltcit1__EACHTUP_ln);
	/*_#NIX__L19*/ meltfnum[0]++)
	{
	  /*_.CURNARG__V41*/ meltfptr[39] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.NARGS__V33*/ meltfptr[31]),
			       /*_#NIX__L19*/ meltfnum[0]);




#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:5989:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L20*/ meltfnum[1] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:5989:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L20*/ meltfnum[1])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L21*/ meltfnum[16] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5989:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L21*/ meltfnum[16];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 5989;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normletrec_list curnarg";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURNARG__V41*/ meltfptr[39];
		    /*_.MELT_DEBUG_FUN__V43*/ meltfptr[42] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V42*/ meltfptr[41] =
		    /*_.MELT_DEBUG_FUN__V43*/ meltfptr[42];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5989:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L21*/ meltfnum[16] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V43*/ meltfptr[42] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V42*/ meltfptr[41] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:5989:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L20*/ meltfnum[1] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V42*/ meltfptr[41] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5990:/ quasiblock");


  /*_.CURCBIND__V45*/ meltfptr[41] =
	    (melt_multiple_nth
	     ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]),
	      ( /*_#NIX__L19*/ meltfnum[0])));;
	  /*^compute */
  /*_#plI__L22*/ meltfnum[17] =
	    ((1) + ( /*_#NIX__L19*/ meltfnum[0]));;
	  /*^compute */
  /*_.NEXTCBIND__V46*/ meltfptr[45] =
	    (melt_multiple_nth
	     ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]),
	      ( /*_#plI__L22*/ meltfnum[17])));;

	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^compute */
  /*_#IS_A__L23*/ meltfnum[16] =
	    melt_is_instance_of ((melt_ptr_t)
				 ( /*_.NEXTCBIND__V46*/ meltfptr[45]),
				 (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])));;
	  MELT_LOCATION ("warmelt-normal.melt:5992:/ cond");
	  /*cond */ if ( /*_#IS_A__L23*/ meltfnum[16])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		MELT_LOCATION ("warmelt-normal.melt:5993:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.NEXTCBIND__V46*/
						     meltfptr[45]),
						    (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[9])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.NEXTCBIND__V46*/ meltfptr[45])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 0, "BINDER");
      /*_.BINDER__V48*/ meltfptr[47] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

     /*_.BINDER__V48*/ meltfptr[47] = NULL;;
		  }
		;
		/*^compute */
		/*_.NEXTPAIRSYMB__V47*/ meltfptr[46] =
		  /*_.BINDER__V48*/ meltfptr[47];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5992:/ clear");
	      /*clear *//*_.BINDER__V48*/ meltfptr[47] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.NEXTPAIRSYMB__V47*/ meltfptr[46] = NULL;;
	    }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:5995:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L24*/ meltfnum[1] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:5995:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L24*/ meltfnum[1])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L25*/ meltfnum[24] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5995:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[9];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L25*/ meltfnum[24];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 5995;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normletrec_list curcbind=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURCBIND__V45*/ meltfptr[41];
		    /*^apply.arg */
		    argtab[5].meltbp_cstring = " nextcbind=";
		    /*^apply.arg */
		    argtab[6].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NEXTCBIND__V46*/ meltfptr[45];
		    /*^apply.arg */
		    argtab[7].meltbp_cstring = " nextpairsymb=";
		    /*^apply.arg */
		    argtab[8].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NEXTPAIRSYMB__V47*/ meltfptr[46];
		    /*_.MELT_DEBUG_FUN__V50*/ meltfptr[49] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V49*/ meltfptr[47] =
		    /*_.MELT_DEBUG_FUN__V50*/ meltfptr[49];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5995:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L25*/ meltfnum[24] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V50*/ meltfptr[49] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V49*/ meltfptr[47] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:5995:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L24*/ meltfnum[1] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V49*/ meltfptr[47] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:5997:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L26*/ meltfnum[24] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURCBIND__V45*/ meltfptr[41]),
				   (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])));;
	    MELT_LOCATION ("warmelt-normal.melt:5997:/ cond");
	    /*cond */ if ( /*_#IS_A__L26*/ meltfnum[24])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V52*/ meltfptr[47] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:5997:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curcbind"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(5997) ? (5997) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V52*/ meltfptr[47] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V51*/ meltfptr[49] =
	      /*_._IFELSE___V52*/ meltfptr[47];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:5997:/ clear");
	      /*clear *//*_#IS_A__L26*/ meltfnum[24] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V52*/ meltfptr[47] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V51*/ meltfptr[49] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:5998:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^cond */
	  /*cond */ if ( /*_.NEXTPAIRSYMB__V47*/ meltfptr[46])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {


		MELT_CHECK_SIGNAL ();
		;
		/*^apply */
		/*apply */
		{
		  union meltparam_un argtab[1];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^apply.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NLETREC__V18*/ meltfptr[17];
		  /*_.NRECLIST_FIND_LOCSYM__V55*/ meltfptr[54] =
		    melt_apply ((meltclosure_ptr_t)
				(( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
				  tabval[10])),
				(melt_ptr_t) ( /*_.NEXTPAIRSYMB__V47*/
					      meltfptr[46]),
				(MELTBPARSTR_PTR ""), argtab, "",
				(union meltparam_un *) 0);
		}
		;
		/*_.NEXTPAIRLOC__V54*/ meltfptr[53] =
		  /*_.NRECLIST_FIND_LOCSYM__V55*/ meltfptr[54];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:5998:/ clear");
	      /*clear *//*_.NRECLIST_FIND_LOCSYM__V55*/ meltfptr[54] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.NEXTPAIRLOC__V54*/ meltfptr[53] = NULL;;
	    }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:5999:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L27*/ meltfnum[1] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:5999:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L27*/ meltfnum[1])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L28*/ meltfnum[24] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:5999:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L28*/ meltfnum[24];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 5999;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normletrec_list got nextpairloc";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NEXTPAIRLOC__V54*/ meltfptr[53];
		    /*_.MELT_DEBUG_FUN__V57*/ meltfptr[56] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V56*/ meltfptr[54] =
		    /*_.MELT_DEBUG_FUN__V57*/ meltfptr[56];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:5999:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L28*/ meltfnum[24] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V57*/ meltfptr[56] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V56*/ meltfptr[54] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:5999:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L27*/ meltfnum[1] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V56*/ meltfptr[54] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6000:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURCBIND__V45*/
					       meltfptr[41]),
					      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		/*^putslot */
		/*putslot */
		melt_assertmsg ("checkobj putslot  _ @NPAIRB_HEAD",
				melt_magic_discr ((melt_ptr_t)
						  ( /*_.CURCBIND__V45*/
						   meltfptr[41])) ==
				MELTOBMAG_OBJECT);
		melt_putfield_object (( /*_.CURCBIND__V45*/ meltfptr[41]),
				      (4), ( /*_.CURNARG__V41*/ meltfptr[39]),
				      "NPAIRB_HEAD");
		;
		/*^putslot */
		/*putslot */
		melt_assertmsg ("checkobj putslot  _ @NPAIRB_TAIL",
				melt_magic_discr ((melt_ptr_t)
						  ( /*_.CURCBIND__V45*/
						   meltfptr[41])) ==
				MELTOBMAG_OBJECT);
		melt_putfield_object (( /*_.CURCBIND__V45*/ meltfptr[41]),
				      (5),
				      ( /*_.NEXTPAIRLOC__V54*/ meltfptr[53]),
				      "NPAIRB_TAIL");
		;
		/*^touch */
		meltgc_touch ( /*_.CURCBIND__V45*/ meltfptr[41]);
		;
		/*^touchobj */

		melt_dbgtrace_written_object ( /*_.CURCBIND__V45*/
					      meltfptr[41], "put-fields");
		;
		/*epilog */
	      }
	      ;
	    }			/*noelse */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6003:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L29*/ meltfnum[24] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6003:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L29*/ meltfnum[24])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L30*/ meltfnum[1] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6003:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L30*/ meltfnum[1];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6003;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normletrec_list updated curcbind";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURCBIND__V45*/ meltfptr[41];
		    /*_.MELT_DEBUG_FUN__V59*/ meltfptr[54] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V58*/ meltfptr[56] =
		    /*_.MELT_DEBUG_FUN__V59*/ meltfptr[54];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6003:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L30*/ meltfnum[1] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V59*/ meltfptr[54] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V58*/ meltfptr[56] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6003:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L29*/ meltfnum[24] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V58*/ meltfptr[56] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6004:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L31*/ meltfnum[1] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURCBIND__V45*/ meltfptr[41]),
				   (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])));;
	    MELT_LOCATION ("warmelt-normal.melt:6004:/ cond");
	    /*cond */ if ( /*_#IS_A__L31*/ meltfnum[1])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V61*/ meltfptr[56] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6004:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curcbind"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6004) ? (6004) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V61*/ meltfptr[56] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V60*/ meltfptr[54] =
	      /*_._IFELSE___V61*/ meltfptr[56];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6004:/ clear");
	      /*clear *//*_#IS_A__L31*/ meltfnum[1] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V61*/ meltfptr[56] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V60*/ meltfptr[54] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  /*^compute */
	  /*_.LET___V53*/ meltfptr[47] = /*_.IFCPP___V60*/ meltfptr[54];;

	  MELT_LOCATION ("warmelt-normal.melt:5998:/ clear");
	    /*clear *//*_.NEXTPAIRLOC__V54*/ meltfptr[53] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V60*/ meltfptr[54] = 0;
	  /*_.LET___V44*/ meltfptr[42] = /*_.LET___V53*/ meltfptr[47];;

	  MELT_LOCATION ("warmelt-normal.melt:5990:/ clear");
	    /*clear *//*_.CURCBIND__V45*/ meltfptr[41] = 0;
	  /*^clear */
	    /*clear *//*_#plI__L22*/ meltfnum[17] = 0;
	  /*^clear */
	    /*clear *//*_.NEXTCBIND__V46*/ meltfptr[45] = 0;
	  /*^clear */
	    /*clear *//*_#IS_A__L23*/ meltfnum[16] = 0;
	  /*^clear */
	    /*clear *//*_.NEXTPAIRSYMB__V47*/ meltfptr[46] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V51*/ meltfptr[49] = 0;
	  /*^clear */
	    /*clear *//*_.LET___V53*/ meltfptr[47] = 0;
	  if ( /*_#NIX__L19*/ meltfnum[0] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:5986:/ clear");
	    /*clear *//*_.CURNARG__V41*/ meltfptr[39] = 0;
      /*^clear */
	    /*clear *//*_#NIX__L19*/ meltfnum[0] = 0;
      /*^clear */
	    /*clear *//*_.LET___V44*/ meltfptr[42] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6006:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L32*/ meltfnum[24] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6006:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L32*/ meltfnum[24])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L33*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6006:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L33*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6006;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list firstpairbind=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.FIRSTPAIRBIND__V21*/ meltfptr[20];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " lastpairbind=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.LASTPAIRBIND__V22*/ meltfptr[21];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " conslibind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSLIBIND__V17*/ meltfptr[16];
	      /*_.MELT_DEBUG_FUN__V63*/ meltfptr[53] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V62*/ meltfptr[56] =
	      /*_.MELT_DEBUG_FUN__V63*/ meltfptr[53];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6006:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L33*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V63*/ meltfptr[53] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V62*/ meltfptr[56] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6006:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L32*/ meltfnum[24] = 0;
      /*^clear */
	     /*clear *//*_._IF___V62*/ meltfptr[56] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#IS_A__L34*/ meltfnum[17] =
      melt_is_instance_of ((melt_ptr_t)
			   ( /*_.FIRSTPAIRBIND__V21*/ meltfptr[20]),
			   (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])));;
    MELT_LOCATION ("warmelt-normal.melt:6009:/ cond");
    /*cond */ if ( /*_#IS_A__L34*/ meltfnum[17])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{


	  MELT_CHECK_SIGNAL ();
	  ;
   /*_#IS_A__L35*/ meltfnum[16] =
	    melt_is_instance_of ((melt_ptr_t)
				 ( /*_.LASTPAIRBIND__V22*/ meltfptr[21]),
				 (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_PAIR_BINDING */ meltfrout->tabval[8])));;
	  MELT_LOCATION ("warmelt-normal.melt:6009:/ cond");
	  /*cond */ if ( /*_#IS_A__L35*/ meltfnum[16])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		MELT_LOCATION ("warmelt-normal.melt:6011:/ quasiblock");


		/*^cond */
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.FIRSTPAIRBIND__V21*/
						     meltfptr[20]),
						    (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[9])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.FIRSTPAIRBIND__V21*/ meltfptr[20])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 0, "BINDER");
       /*_.FIRSTPAIRSYMB__V64*/ meltfptr[54] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

      /*_.FIRSTPAIRSYMB__V64*/ meltfptr[54] = NULL;;
		  }
		;
		MELT_LOCATION ("warmelt-normal.melt:6012:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.LASTPAIRBIND__V22*/
						     meltfptr[21]),
						    (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[9])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.LASTPAIRBIND__V22*/ meltfptr[21])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 0, "BINDER");
       /*_.LASTPAIRSYMB__V65*/ meltfptr[41] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

      /*_.LASTPAIRSYMB__V65*/ meltfptr[41] = NULL;;
		  }
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6013:/ apply");
		/*apply */
		{
		  union meltparam_un argtab[1];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^apply.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NLETREC__V18*/ meltfptr[17];
		  /*_.FIRSTPAIRLOCSY__V66*/ meltfptr[45] =
		    melt_apply ((meltclosure_ptr_t)
				(( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
				  tabval[10])),
				(melt_ptr_t) ( /*_.FIRSTPAIRSYMB__V64*/
					      meltfptr[54]),
				(MELTBPARSTR_PTR ""), argtab, "",
				(union meltparam_un *) 0);
		}
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6014:/ apply");
		/*apply */
		{
		  union meltparam_un argtab[1];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^apply.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NLETREC__V18*/ meltfptr[17];
		  /*_.LASTPAIRLOCSY__V67*/ meltfptr[46] =
		    melt_apply ((meltclosure_ptr_t)
				(( /*!NRECLIST_FIND_LOCSYM */ meltfrout->
				  tabval[10])),
				(melt_ptr_t) ( /*_.LASTPAIRSYMB__V65*/
					      meltfptr[41]),
				(MELTBPARSTR_PTR ""), argtab, "",
				(union meltparam_un *) 0);
		}
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6016:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
       /*_#MELT_NEED_DBG__L36*/ meltfnum[1] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:6016:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L36*/ meltfnum[1])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	 /*_#MELT_CALLCOUNT__L37*/ meltfnum[24] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:6016:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[7];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L37*/ meltfnum[24];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 6016;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "normletrec_list firstpairlocsy=";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.FIRSTPAIRLOCSY__V66*/
			    meltfptr[45];
			  /*^apply.arg */
			  argtab[5].meltbp_cstring = " lastpairlocsy=";
			  /*^apply.arg */
			  argtab[6].meltbp_aptr =
			    (melt_ptr_t *) & /*_.LASTPAIRLOCSY__V67*/
			    meltfptr[46];
			  /*_.MELT_DEBUG_FUN__V69*/ meltfptr[47] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[0])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V68*/ meltfptr[49] =
			  /*_.MELT_DEBUG_FUN__V69*/ meltfptr[47];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6016:/ clear");
		   /*clear *//*_#MELT_CALLCOUNT__L37*/ meltfnum[24] = 0;
			/*^clear */
		   /*clear *//*_.MELT_DEBUG_FUN__V69*/ meltfptr[47] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

	/*_._IF___V68*/ meltfptr[49] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:6016:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		 /*clear *//*_#MELT_NEED_DBG__L36*/ meltfnum[1] = 0;
		  /*^clear */
		 /*clear *//*_._IF___V68*/ meltfptr[49] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6017:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.CONSLIBIND__V17*/
						     meltfptr[16]),
						    (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_LIST_BINDING */ meltfrout->tabval[6])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^block */
		    /*anyblock */
		    {

		      /*^putslot */
		      /*putslot */
		      melt_assertmsg ("checkobj putslot  _ @NLISTB_FIRST",
				      melt_magic_discr ((melt_ptr_t)
							( /*_.CONSLIBIND__V17*/ meltfptr[16])) == MELTOBMAG_OBJECT);
		      melt_putfield_object (( /*_.CONSLIBIND__V17*/
					     meltfptr[16]), (4),
					    ( /*_.FIRSTPAIRLOCSY__V66*/
					     meltfptr[45]), "NLISTB_FIRST");
		      ;
		      /*^putslot */
		      /*putslot */
		      melt_assertmsg ("checkobj putslot  _ @NLISTB_LAST",
				      melt_magic_discr ((melt_ptr_t)
							( /*_.CONSLIBIND__V17*/ meltfptr[16])) == MELTOBMAG_OBJECT);
		      melt_putfield_object (( /*_.CONSLIBIND__V17*/
					     meltfptr[16]), (5),
					    ( /*_.LASTPAIRLOCSY__V67*/
					     meltfptr[46]), "NLISTB_LAST");
		      ;
		      /*^touch */
		      meltgc_touch ( /*_.CONSLIBIND__V17*/ meltfptr[16]);
		      ;
		      /*^touchobj */

		      melt_dbgtrace_written_object ( /*_.CONSLIBIND__V17*/
						    meltfptr[16],
						    "put-fields");
		      ;
		      /*epilog */
		    }
		    ;
		  }		/*noelse */
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6020:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
       /*_#MELT_NEED_DBG__L38*/ meltfnum[24] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:6020:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L38*/ meltfnum[24])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	 /*_#MELT_CALLCOUNT__L39*/ meltfnum[1] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:6020:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[5];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L39*/ meltfnum[1];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 6020;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "normletrec_list updated conslibind";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CONSLIBIND__V17*/
			    meltfptr[16];
			  /*_.MELT_DEBUG_FUN__V71*/ meltfptr[56] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[0])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V70*/ meltfptr[53] =
			  /*_.MELT_DEBUG_FUN__V71*/ meltfptr[56];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6020:/ clear");
		   /*clear *//*_#MELT_CALLCOUNT__L39*/ meltfnum[1] = 0;
			/*^clear */
		   /*clear *//*_.MELT_DEBUG_FUN__V71*/ meltfptr[56] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

	/*_._IF___V70*/ meltfptr[53] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:6020:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		 /*clear *//*_#MELT_NEED_DBG__L38*/ meltfnum[24] = 0;
		  /*^clear */
		 /*clear *//*_._IF___V70*/ meltfptr[53] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

		{
		  MELT_LOCATION ("warmelt-normal.melt:6021:/ locexp");
		  /*void */ (void) 0;
		}
		;

		MELT_LOCATION ("warmelt-normal.melt:6011:/ clear");
	       /*clear *//*_.FIRSTPAIRSYMB__V64*/ meltfptr[54] = 0;
		/*^clear */
	       /*clear *//*_.LASTPAIRSYMB__V65*/ meltfptr[41] = 0;
		/*^clear */
	       /*clear *//*_.FIRSTPAIRLOCSY__V66*/ meltfptr[45] = 0;
		/*^clear */
	       /*clear *//*_.LASTPAIRLOCSY__V67*/ meltfptr[46] = 0;
		/*epilog */
	      }
	      ;
	    }			/*noelse */
	  ;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6009:/ clear");
	     /*clear *//*_#IS_A__L35*/ meltfnum[16] = 0;
	}
	;
      }				/*noelse */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6023:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L40*/ meltfnum[1] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6023:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L40*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L41*/ meltfnum[24] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6023:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L41*/ meltfnum[24];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6023;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_list appending nbinds";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBINDS__V34*/ meltfptr[33];
	      /*_.MELT_DEBUG_FUN__V73*/ meltfptr[49] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V72*/ meltfptr[47] =
	      /*_.MELT_DEBUG_FUN__V73*/ meltfptr[49];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6023:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L41*/ meltfnum[24] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V73*/ meltfptr[49] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V72*/ meltfptr[47] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6023:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L40*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IF___V72*/ meltfptr[47] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6024:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBINDS__V34*/ meltfptr[33];
      /*_.LIST_APPEND2LIST__V74*/ meltfptr[56] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_APPEND2LIST */ meltfrout->tabval[11])),
		    (melt_ptr_t) ( /*_.NRECBINDS__V19*/ meltfptr[18]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6025:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L42*/ meltfnum[16] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6025:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L42*/ meltfnum[16])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L43*/ meltfnum[24] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6025:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L43*/ meltfnum[24];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6025;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normletrec_list ended updated nrecbinds";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NRECBINDS__V19*/ meltfptr[18];
	      /*_.MELT_DEBUG_FUN__V76*/ meltfptr[54] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V75*/ meltfptr[53] =
	      /*_.MELT_DEBUG_FUN__V76*/ meltfptr[54];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6025:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L43*/ meltfnum[24] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V76*/ meltfptr[54] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V75*/ meltfptr[53] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6025:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L42*/ meltfnum[16] = 0;
      /*^clear */
	     /*clear *//*_._IF___V75*/ meltfptr[53] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;

    MELT_LOCATION ("warmelt-normal.melt:5979:/ clear");
	   /*clear *//*_.IFCPP___V37*/ meltfptr[35] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V39*/ meltfptr[34] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L34*/ meltfnum[17] = 0;
    /*^clear */
	   /*clear *//*_.LIST_APPEND2LIST__V74*/ meltfptr[56] = 0;

    /*^clear */
	   /*clear *//*_.NBINDS__V34*/ meltfptr[33] = 0;

    MELT_LOCATION ("warmelt-normal.melt:5963:/ clear");
	   /*clear *//*_.LOCA_LOCATION__V14*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.LOC__V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.SARGS__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.CONSLIBIND__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.NLETREC__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.NRECBINDS__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.NLOCSYMS__V20*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_.FIRSTPAIRBIND__V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.LASTPAIRBIND__V22*/ meltfptr[21] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V25*/ meltfptr[23] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V29*/ meltfptr[27] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V31*/ meltfptr[22] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */

    MELT_LOCATION ("warmelt-normal.melt:5957:/ clear");
	   /*clear *//*_.IFCPP___V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V12*/ meltfptr[7] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMLETREC_LIST", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*noretval */ NULL);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_110_WARMELTmiNORMAL_NORMLETREC_LIST */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un *
						meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un *
						meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 25
    melt_ptr_t mcfr_varptr[25];
#define MELTFRAM_NBVARNUM 8
    long mcfr_varnum[8];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE_st *)
	meltfirstargp_;
      /* use arguments meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 25; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE nbval 25*/
  meltfram__.mcfr_nbvar = 25 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("PREPCONS_INSTANCE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6031:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6032:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6032:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6032:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6032;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_instance recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6032:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6032:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6033:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_INSTANCE */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6033:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6033:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6033) ? (6033) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6033:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6034:/ quasiblock");


    MELT_LOCATION ("warmelt-normal.melt:6035:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_LOCATED */
						       meltfrout->tabval[2])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOC__V11*/ meltfptr[10] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.LOC__V11*/ meltfptr[10] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6036:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_SOURCE_INSTANCE */ meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 2, "SMINS_CLASS");
   /*_.CLA__V12*/ meltfptr[11] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.CLA__V12*/ meltfptr[11] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6037:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.RECV__V2*/ meltfptr[1]),
					(melt_ptr_t) (( /*!CLASS_SOURCE_INSTANCE */ meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "SMINS_CLABIND");
   /*_.CLABIND__V13*/ meltfptr[12] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.CLABIND__V13*/ meltfptr[12] = NULL;;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    /*^compute */
 /*_#IS_A__L4*/ meltfnum[0] =
      melt_is_instance_of ((melt_ptr_t) ( /*_.CLABIND__V13*/ meltfptr[12]),
			   (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->
					  tabval[3])));;
    MELT_LOCATION ("warmelt-normal.melt:6038:/ cond");
    /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  /*^getslot */
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CLABIND__V13*/ meltfptr[12]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 0, "BINDER");
    /*_.BINDER__V15*/ meltfptr[14] = slot;
	  };
	  ;
	  /*_.SCLASYM__V14*/ meltfptr[13] = /*_.BINDER__V15*/ meltfptr[14];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6038:/ clear");
	     /*clear *//*_.BINDER__V15*/ meltfptr[14] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

  /*_.SCLASYM__V14*/ meltfptr[13] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6039:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CLA__V12*/ meltfptr[11]),
					(melt_ptr_t) (( /*!CLASS_CLASS */
						       meltfrout->tabval[4])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CLA__V12*/ meltfptr[11]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 6, "CLASS_FIELDS");
   /*_.CLASS_FIELDS__V16*/ meltfptr[14] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.CLASS_FIELDS__V16*/ meltfptr[14] = NULL;;
      }
    ;
    /*^compute */
 /*_#NBCLAFLD__L5*/ meltfnum[1] =
      (melt_multiple_length
       ((melt_ptr_t) ( /*_.CLASS_FIELDS__V16*/ meltfptr[14])));;
    /*^compute */
 /*_.TUPSLOT__V17*/ meltfptr[16] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[5])),
	( /*_#NBCLAFLD__L5*/ meltfnum[1])));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6043:/ cond");
    /*cond */ if ( /*_.LOC__V11*/ meltfptr[10])	/*then */
      {
	/*^cond.then */
	/*_.OR___V18*/ meltfptr[17] = /*_.LOC__V11*/ meltfptr[10];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6043:/ cond.else");

	/*_.OR___V18*/ meltfptr[17] = /*_.SLOC__V5*/ meltfptr[4];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6041:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_INSTANCE_BINDING */ meltfrout->tabval[6])), (6), "CLASS_NORMAL_CONSTRUCTED_INSTANCE_BINDING");
  /*_.INST__V20*/ meltfptr[19] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @BINDER",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (0),
			  ( /*_.SYMB__V3*/ meltfptr[2]), "BINDER");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (1),
			  ( /*_.OR___V18*/ meltfptr[17]), "NCONSB_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (2),
			  (( /*nil */ NULL)), "NCONSB_DISCR");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NINSTB_SLOTS",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (4),
			  ( /*_.TUPSLOT__V17*/ meltfptr[16]), "NINSTB_SLOTS");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NINSTB_CLABIND",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (5),
			  ( /*_.CLABIND__V13*/ meltfptr[12]),
			  "NINSTB_CLABIND");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V20*/ meltfptr[19],
				  "newly made instance");
    ;
    /*_.CONSINST__V19*/ meltfptr[18] = /*_.INST__V20*/ meltfptr[19];;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6049:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L6*/ meltfnum[5] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CLA__V12*/ meltfptr[11]),
			     (melt_ptr_t) (( /*!CLASS_CLASS */ meltfrout->
					    tabval[4])));;
      MELT_LOCATION ("warmelt-normal.melt:6049:/ cond");
      /*cond */ if ( /*_#IS_A__L6*/ meltfnum[5])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V22*/ meltfptr[21] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6049:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("prepcons_instance check class"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6049) ? (6049) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V22*/ meltfptr[21] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V21*/ meltfptr[20] = /*_._IFELSE___V22*/ meltfptr[21];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6049:/ clear");
	     /*clear *//*_#IS_A__L6*/ meltfnum[5] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V22*/ meltfptr[21] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V21*/ meltfptr[20] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6050:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L7*/ meltfnum[5] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6050:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L7*/ meltfnum[5])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L8*/ meltfnum[7] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6050:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L8*/ meltfnum[7];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6050;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "prepcons_instance gives consinst";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSINST__V19*/ meltfptr[18];
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V23*/ meltfptr[21] =
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6050:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L8*/ meltfnum[7] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V23*/ meltfptr[21] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6050:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L7*/ meltfnum[5] = 0;
      /*^clear */
	     /*clear *//*_._IF___V23*/ meltfptr[21] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6051:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.CONSINST__V19*/ meltfptr[18];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6051:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V10*/ meltfptr[5] = /*_.RETURN___V25*/ meltfptr[23];;

    MELT_LOCATION ("warmelt-normal.melt:6034:/ clear");
	   /*clear *//*_.LOC__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.CLA__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_.CLABIND__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.SCLASYM__V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_.CLASS_FIELDS__V16*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_#NBCLAFLD__L5*/ meltfnum[1] = 0;
    /*^clear */
	   /*clear *//*_.TUPSLOT__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.OR___V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.CONSINST__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V25*/ meltfptr[23] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6031:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V10*/ meltfptr[5];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6031:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.LET___V10*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("PREPCONS_INSTANCE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_111_WARMELTmiNORMAL_PREPCONS_INSTANCE */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un *
						  meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un *
						  meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 46
    melt_ptr_t mcfr_varptr[46];
#define MELTFRAM_NBVARNUM 18
    long mcfr_varnum[18];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE_st
	 *) meltfirstargp_;
      /* use arguments meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 46; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE nbval 46*/
  meltfram__.mcfr_nbvar = 46 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMLETREC_INSTANCE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6058:/ getarg");
 /*_.CEXPR__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SYMB__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SYMB__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.CBIND__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4])) != NULL);


  /*getarg#4 */
  /*^getarg */
  if (meltxargdescr_[3] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V6*/ meltfptr[5] =
    (meltxargtab_[3].meltbp_aptr) ? (*(meltxargtab_[3].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5])) != NULL);


  /*getarg#5 */
  /*^getarg */
  if (meltxargdescr_[4] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V7*/ meltfptr[6] =
    (meltxargtab_[4].meltbp_aptr) ? (*(meltxargtab_[4].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6059:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6059:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6059:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6059;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_instance cexpr=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CEXPR__V2*/ meltfptr[1];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " symb=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYMB__V3*/ meltfptr[2];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " cbind=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.CBIND__V4*/ meltfptr[3];
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V8*/ meltfptr[7] =
	      /*_.MELT_DEBUG_FUN__V9*/ meltfptr[8];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6059:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V9*/ meltfptr[8] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V8*/ meltfptr[7] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6059:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6060:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.LOCA_LOCATION__V10*/ meltfptr[8] = slot;
    };
    ;
    /*^cond */
    /*cond */ if ( /*_.LOCA_LOCATION__V10*/ meltfptr[8])	/*then */
      {
	/*^cond.then */
	/*_.LOC__V11*/ meltfptr[7] = /*_.LOCA_LOCATION__V10*/ meltfptr[8];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6060:/ cond.else");

	/*_.LOC__V11*/ meltfptr[7] = /*_.SLOC__V7*/ meltfptr[6];;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6061:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CBIND__V4*/ meltfptr[3]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_INSTANCE_BINDING */ meltfrout->tabval[1])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 4, "NINSTB_SLOTS");
   /*_.BSLOTS__V12*/ meltfptr[11] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.BSLOTS__V12*/ meltfptr[11] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6062:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.CBIND__V4*/ meltfptr[3]),
					(melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[2])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.CBIND__V4*/ meltfptr[3]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NCONSB_NLETREC");
   /*_.NLETREC__V13*/ meltfptr[12] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NLETREC__V13*/ meltfptr[12] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6063:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V13*/ meltfptr[12]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[3])))
      )				/*then */
      {
	/*^cond.then */
	/*^getslot */
	{
	  melt_ptr_t slot = NULL, obj = NULL;
	  obj = (melt_ptr_t) ( /*_.NLETREC__V13*/ meltfptr[12]) /*=obj*/ ;
	  melt_object_get_field (slot, obj, 3, "NLETREC_FILL_BINDINGS");
   /*_.NRECBINDS__V14*/ meltfptr[13] = slot;
	};
	;
      }
    else
      {				/*^cond.else */

  /*_.NRECBINDS__V14*/ meltfptr[13] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6064:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SMINS_CLASS");
  /*_.SCLASS__V15*/ meltfptr[14] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6065:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 3, "SMINS_CLABIND");
  /*_.SCLABIND__V16*/ meltfptr[15] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6066:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CEXPR__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 4, "SMINS_FIELDS");
  /*_.SFIELDS__V17*/ meltfptr[16] = slot;
    };
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#IS_A__L3*/ meltfnum[1] =
      melt_is_instance_of ((melt_ptr_t) ( /*_.SCLABIND__V16*/ meltfptr[15]),
			   (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->
					  tabval[4])));;
    MELT_LOCATION ("warmelt-normal.melt:6067:/ cond");
    /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  /*^getslot */
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.SCLABIND__V16*/ meltfptr[15]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 0, "BINDER");
    /*_.BINDER__V19*/ meltfptr[18] = slot;
	  };
	  ;
	  /*_.SCLASYM__V18*/ meltfptr[17] = /*_.BINDER__V19*/ meltfptr[18];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6067:/ clear");
	     /*clear *//*_.BINDER__V19*/ meltfptr[18] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

  /*_.SCLASYM__V18*/ meltfptr[17] = NULL;;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    /*^compute */
 /*_#IS_A__L4*/ meltfnum[0] =
      melt_is_instance_of ((melt_ptr_t) ( /*_.SCLASYM__V18*/ meltfptr[17]),
			   (melt_ptr_t) (( /*!CLASS_SYMBOL */ meltfrout->
					  tabval[5])));;
    MELT_LOCATION ("warmelt-normal.melt:6068:/ cond");
    /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{


	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6069:/ msend");
	  /*msend */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];
	    /*^ojbmsend.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V6*/ meltfptr[5];
	    /*^ojbmsend.arg */
	    argtab[2].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V7*/ meltfptr[6];
	    /*_.NORMAL_EXP__V21*/ meltfptr[20] =
	      meltgc_send ((melt_ptr_t) ( /*_.SCLASYM__V18*/ meltfptr[17]),
			   (melt_ptr_t) (( /*!NORMAL_EXP */ meltfrout->
					  tabval[6])),
			   (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
			    ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
	  /*_.CLADATA__V20*/ meltfptr[18] =
	    /*_.NORMAL_EXP__V21*/ meltfptr[20];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6068:/ clear");
	     /*clear *//*_.NORMAL_EXP__V21*/ meltfptr[20] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

  /*_.CLADATA__V20*/ meltfptr[18] = NULL;;
      }
    ;
    /*^compute */
 /*_.BINDLIST__V22*/ meltfptr[20] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[7]))));;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6072:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L5*/ meltfnum[4] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6072:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L5*/ meltfnum[4])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L6*/ meltfnum[5] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6072:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L6*/ meltfnum[5];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6072;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_instance nletrec";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V13*/ meltfptr[12];
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V23*/ meltfptr[22] =
	      /*_.MELT_DEBUG_FUN__V24*/ meltfptr[23];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6072:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L6*/ meltfnum[5] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V24*/ meltfptr[23] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V23*/ meltfptr[22] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6072:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L5*/ meltfnum[4] = 0;
      /*^clear */
	     /*clear *//*_._IF___V23*/ meltfptr[22] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#IS_A__L7*/ meltfnum[5] =
      melt_is_instance_of ((melt_ptr_t) ( /*_.CLADATA__V20*/ meltfptr[18]),
			   (melt_ptr_t) (( /*!CLASS_NREP */ meltfrout->
					  tabval[8])));;
    /*^compute */
 /*_#NOT__L8*/ meltfnum[4] =
      (!( /*_#IS_A__L7*/ meltfnum[5]));;
    MELT_LOCATION ("warmelt-normal.melt:6073:/ cond");
    /*cond */ if ( /*_#NOT__L8*/ meltfnum[4])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:6075:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.SCLASS__V15*/ meltfptr[14]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 1, "NAMED_NAME");
    /*_.NAMED_NAME__V26*/ meltfptr[22] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    melt_error_str ((melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6]),
			    ("invalid class in (INSTANCE ..)"),
			    (melt_ptr_t) ( /*_.NAMED_NAME__V26*/
					  meltfptr[22]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6076:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = ( /*nil */ NULL);;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6076:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  MELT_LOCATION ("warmelt-normal.melt:6074:/ quasiblock");


	  /*_.PROGN___V28*/ meltfptr[27] = /*_.RETURN___V27*/ meltfptr[26];;
	  /*^compute */
	  /*_._IF___V25*/ meltfptr[23] = /*_.PROGN___V28*/ meltfptr[27];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6073:/ clear");
	     /*clear *//*_.NAMED_NAME__V26*/ meltfptr[22] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V27*/ meltfptr[26] = 0;
	  /*^clear */
	     /*clear *//*_.PROGN___V28*/ meltfptr[27] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

  /*_._IF___V25*/ meltfptr[23] = NULL;;
      }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6077:/ quasiblock");


    MELT_LOCATION ("warmelt-normal.melt:6080:/ quasiblock");


    /*^newclosure */
		 /*newclosure *//*_.LAMBDA___V30*/ meltfptr[26] =
      (melt_ptr_t)
      meltgc_new_closure ((meltobject_ptr_t)
			  (((melt_ptr_t) (MELT_PREDEF (DISCR_CLOSURE)))),
			  (meltroutine_ptr_t) (( /*!konst_15 */ meltfrout->
						tabval[15])), (4));
    ;
    /*^putclosedv */
    /*putclosv */
    melt_assertmsg ("putclosv checkclo",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.LAMBDA___V30*/ meltfptr[26])) ==
		    MELTOBMAG_CLOSURE);
    melt_assertmsg ("putclosv checkoff", 0 >= 0
		    && 0 <
		    melt_closure_size ((melt_ptr_t)
				       ( /*_.LAMBDA___V30*/ meltfptr[26])));
    ((meltclosure_ptr_t) /*_.LAMBDA___V30*/ meltfptr[26])->tabval[0] =
      (melt_ptr_t) ( /*_.SLOC__V7*/ meltfptr[6]);
    ;
    /*^putclosedv */
    /*putclosv */
    melt_assertmsg ("putclosv checkclo",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.LAMBDA___V30*/ meltfptr[26])) ==
		    MELTOBMAG_CLOSURE);
    melt_assertmsg ("putclosv checkoff", 1 >= 0
		    && 1 <
		    melt_closure_size ((melt_ptr_t)
				       ( /*_.LAMBDA___V30*/ meltfptr[26])));
    ((meltclosure_ptr_t) /*_.LAMBDA___V30*/ meltfptr[26])->tabval[1] =
      (melt_ptr_t) ( /*_.ENV__V5*/ meltfptr[4]);
    ;
    /*^putclosedv */
    /*putclosv */
    melt_assertmsg ("putclosv checkclo",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.LAMBDA___V30*/ meltfptr[26])) ==
		    MELTOBMAG_CLOSURE);
    melt_assertmsg ("putclosv checkoff", 2 >= 0
		    && 2 <
		    melt_closure_size ((melt_ptr_t)
				       ( /*_.LAMBDA___V30*/ meltfptr[26])));
    ((meltclosure_ptr_t) /*_.LAMBDA___V30*/ meltfptr[26])->tabval[2] =
      (melt_ptr_t) ( /*_.NCX__V6*/ meltfptr[5]);
    ;
    /*^putclosedv */
    /*putclosv */
    melt_assertmsg ("putclosv checkclo",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.LAMBDA___V30*/ meltfptr[26])) ==
		    MELTOBMAG_CLOSURE);
    melt_assertmsg ("putclosv checkoff", 3 >= 0
		    && 3 <
		    melt_closure_size ((melt_ptr_t)
				       ( /*_.LAMBDA___V30*/ meltfptr[26])));
    ((meltclosure_ptr_t) /*_.LAMBDA___V30*/ meltfptr[26])->tabval[3] =
      (melt_ptr_t) ( /*_.BINDLIST__V22*/ meltfptr[20]);
    ;
    /*_.LAMBDA___V29*/ meltfptr[22] = /*_.LAMBDA___V30*/ meltfptr[26];;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6078:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.LAMBDA___V29*/ meltfptr[22];
      /*_.NFIELDS__V31*/ meltfptr[27] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!MULTIPLE_MAP */ meltfrout->tabval[9])),
		    (melt_ptr_t) ( /*_.SFIELDS__V17*/ meltfptr[16]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6098:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L9*/ meltfnum[8] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6098:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L9*/ meltfnum[8])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L10*/ meltfnum[9] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6098:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L10*/ meltfnum[9];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6098;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_instance nfields";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NFIELDS__V31*/ meltfptr[27];
	      /*_.MELT_DEBUG_FUN__V33*/ meltfptr[32] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V32*/ meltfptr[31] =
	      /*_.MELT_DEBUG_FUN__V33*/ meltfptr[32];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6098:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L10*/ meltfnum[9] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V33*/ meltfptr[32] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V32*/ meltfptr[31] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6098:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L9*/ meltfnum[8] = 0;
      /*^clear */
	     /*clear *//*_._IF___V32*/ meltfptr[31] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.NFIELDS__V31*/ meltfptr[27]);
      for ( /*_#FLDIX__L11*/ meltfnum[9] = 0;
	   ( /*_#FLDIX__L11*/ meltfnum[9] >= 0)
	   && ( /*_#FLDIX__L11*/ meltfnum[9] < meltcit1__EACHTUP_ln);
	/*_#FLDIX__L11*/ meltfnum[9]++)
	{
	  /*_.CURNFIELDASS__V34*/ meltfptr[32] =
	    melt_multiple_nth ((melt_ptr_t)
			       ( /*_.NFIELDS__V31*/ meltfptr[27]),
			       /*_#FLDIX__L11*/ meltfnum[9]);




#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6102:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L12*/ meltfnum[8] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6102:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L12*/ meltfnum[8])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L13*/ meltfnum[12] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6102:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L13*/ meltfnum[12];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6102;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normletrec_instance curnfieldass";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURNFIELDASS__V34*/ meltfptr[32];
		    /*_.MELT_DEBUG_FUN__V36*/ meltfptr[35] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V35*/ meltfptr[31] =
		    /*_.MELT_DEBUG_FUN__V36*/ meltfptr[35];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6102:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L13*/ meltfnum[12] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V36*/ meltfptr[35] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V35*/ meltfptr[31] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6102:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L12*/ meltfnum[8] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V35*/ meltfptr[31] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6103:/ quasiblock");


	  /*^cond */
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURNFIELDASS__V34*/
					       meltfptr[32]),
					      (melt_ptr_t) (( /*!CLASS_NREP_FIELDASSIGN */ meltfrout->tabval[16])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURNFIELDASS__V34*/ meltfptr[32])
		  /*=obj*/ ;
		melt_object_get_field (slot, obj, 1, "NFLA_FIELD");
    /*_.CURFIELD__V37*/ meltfptr[35] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.CURFIELD__V37*/ meltfptr[35] = NULL;;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6104:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURNFIELDASS__V34*/
					       meltfptr[32]),
					      (melt_ptr_t) (( /*!CLASS_NREP_FIELDASSIGN */ meltfrout->tabval[16])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURNFIELDASS__V34*/ meltfptr[32])
		  /*=obj*/ ;
		melt_object_get_field (slot, obj, 2, "NFLA_VAL");
    /*_.CURFVAL__V38*/ meltfptr[31] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.CURFVAL__V38*/ meltfptr[31] = NULL;;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6105:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURNFIELDASS__V34*/
					       meltfptr[32]),
					      (melt_ptr_t) (( /*!CLASS_NREP */
							     meltfrout->
							     tabval[8])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURNFIELDASS__V34*/ meltfptr[32])
		  /*=obj*/ ;
		melt_object_get_field (slot, obj, 0, "NREP_LOC");
    /*_.CURFLOC__V39*/ meltfptr[38] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.CURFLOC__V39*/ meltfptr[38] = NULL;;
	    }
	  ;
	  /*^compute */
  /*_#CURFOFF__L14*/ meltfnum[12] =
	    (melt_get_int
	     ((melt_ptr_t) ( /*_.CURFIELD__V37*/ meltfptr[35])));;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6108:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L15*/ meltfnum[8] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURFIELD__V37*/ meltfptr[35]),
				   (melt_ptr_t) (( /*!CLASS_FIELD */
						  meltfrout->tabval[17])));;
	    MELT_LOCATION ("warmelt-normal.melt:6108:/ cond");
	    /*cond */ if ( /*_#IS_A__L15*/ meltfnum[8])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V41*/ meltfptr[40] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6108:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curfield "),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6108) ? (6108) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V41*/ meltfptr[40] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V40*/ meltfptr[39] =
	      /*_._IFELSE___V41*/ meltfptr[40];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6108:/ clear");
	      /*clear *//*_#IS_A__L15*/ meltfnum[8] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V41*/ meltfptr[40] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V40*/ meltfptr[39] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6109:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^msend */
	  /*msend */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.ENV__V5*/ meltfptr[4];
	    /*_.FCTYP__V42*/ meltfptr[40] =
	      meltgc_send ((melt_ptr_t) ( /*_.CURFVAL__V38*/ meltfptr[31]),
			   (melt_ptr_t) (( /*!GET_CTYPE */ meltfrout->
					  tabval[18])), (MELTBPARSTR_PTR ""),
			   argtab, "", (union meltparam_un *) 0);
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
  /*_#exeq__L16*/ meltfnum[8] =
	    (( /*_.FCTYP__V42*/ meltfptr[40]) !=
	     (( /*!CTYPE_VALUE */ meltfrout->tabval[19])));;
	  MELT_LOCATION ("warmelt-normal.melt:6111:/ cond");
	  /*cond */ if ( /*_#exeq__L16*/ meltfnum[8])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		MELT_LOCATION ("warmelt-normal.melt:6115:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.CURFIELD__V37*/
						     meltfptr[35]),
						    (melt_ptr_t) (( /*!CLASS_NAMED */ meltfrout->tabval[20])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.CURFIELD__V37*/ meltfptr[35])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 1, "NAMED_NAME");
      /*_.NAMED_NAME__V43*/ meltfptr[42] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

     /*_.NAMED_NAME__V43*/ meltfptr[42] = NULL;;
		  }
		;

		{
		  MELT_LOCATION ("warmelt-normal.melt:6112:/ locexp");
		  melt_error_str ((melt_ptr_t)
				  ( /*_.CURFLOC__V39*/ meltfptr[38]),
				  ("invalid field type in (LETREC .. (INSTANCE ..); expecting a :value"),
				  (melt_ptr_t) ( /*_.NAMED_NAME__V43*/
						meltfptr[42]));
		}
		;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6111:/ clear");
	      /*clear *//*_.NAMED_NAME__V43*/ meltfptr[42] = 0;
	      }
	      ;
	    }			/*noelse */
	  ;

	  MELT_LOCATION ("warmelt-normal.melt:6109:/ clear");
	    /*clear *//*_.FCTYP__V42*/ meltfptr[40] = 0;
	  /*^clear */
	    /*clear *//*_#exeq__L16*/ meltfnum[8] = 0;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6116:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.BSLOTS__V12*/ meltfptr[11]),
				     ( /*_#CURFOFF__L14*/ meltfnum[12]),
				     (melt_ptr_t) ( /*_.CURFVAL__V38*/
						   meltfptr[31]));
	  }
	  ;

	  MELT_LOCATION ("warmelt-normal.melt:6103:/ clear");
	    /*clear *//*_.CURFIELD__V37*/ meltfptr[35] = 0;
	  /*^clear */
	    /*clear *//*_.CURFVAL__V38*/ meltfptr[31] = 0;
	  /*^clear */
	    /*clear *//*_.CURFLOC__V39*/ meltfptr[38] = 0;
	  /*^clear */
	    /*clear *//*_#CURFOFF__L14*/ meltfnum[12] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V40*/ meltfptr[39] = 0;
	  if ( /*_#FLDIX__L11*/ meltfnum[9] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:6099:/ clear");
	    /*clear *//*_.CURNFIELDASS__V34*/ meltfptr[32] = 0;
      /*^clear */
	    /*clear *//*_#FLDIX__L11*/ meltfnum[9] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6118:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.BINDLIST__V22*/ meltfptr[20];
      /*_.LIST_APPEND2LIST__V44*/ meltfptr[42] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_APPEND2LIST */ meltfrout->tabval[21])),
		    (melt_ptr_t) ( /*_.NRECBINDS__V14*/ meltfptr[13]),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6119:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L17*/ meltfnum[8] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6119:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L17*/ meltfnum[8])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L18*/ meltfnum[12] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6119:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L18*/ meltfnum[12];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6119;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normletrec_instance ended updated nrecbinds";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NRECBINDS__V14*/ meltfptr[13];
	      /*_.MELT_DEBUG_FUN__V46*/ meltfptr[35] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V45*/ meltfptr[40] =
	      /*_.MELT_DEBUG_FUN__V46*/ meltfptr[35];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6119:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L18*/ meltfnum[12] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V46*/ meltfptr[35] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V45*/ meltfptr[40] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6119:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L17*/ meltfnum[8] = 0;
      /*^clear */
	     /*clear *//*_._IF___V45*/ meltfptr[40] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_LOCATION ("warmelt-normal.melt:6077:/ clear");
	   /*clear *//*_.LAMBDA___V29*/ meltfptr[22] = 0;
    /*^clear */
	   /*clear *//*_.NFIELDS__V31*/ meltfptr[27] = 0;
    /*^clear */
	   /*clear *//*_.LIST_APPEND2LIST__V44*/ meltfptr[42] = 0;

    MELT_LOCATION ("warmelt-normal.melt:6060:/ clear");
	   /*clear *//*_.LOCA_LOCATION__V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.LOC__V11*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_.BSLOTS__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_.NLETREC__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.NRECBINDS__V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_.SCLASS__V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.SCLABIND__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.SFIELDS__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
    /*^clear */
	   /*clear *//*_.SCLASYM__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.CLADATA__V20*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.BINDLIST__V22*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L7*/ meltfnum[5] = 0;
    /*^clear */
	   /*clear *//*_#NOT__L8*/ meltfnum[4] = 0;
    /*^clear */
	   /*clear *//*_._IF___V25*/ meltfptr[23] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMLETREC_INSTANCE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_112_WARMELTmiNORMAL_NORMLETREC_INSTANCE */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 (meltclosure_ptr_t meltclosp_,
					  melt_ptr_t meltfirstargp_,
					  const melt_argdescr_cell_t
					  meltxargdescr_[],
					  union meltparam_un * meltxargtab_,
					  const melt_argdescr_cell_t
					  meltxresdescr_[],
					  union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33_melt =
    melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 20
    melt_ptr_t mcfr_varptr[20];
#define MELTFRAM_NBVARNUM 6
    long mcfr_varnum[6];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33_st *)
	meltfirstargp_;
      /* use arguments meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 20; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 nbval 20*/
  meltfram__.mcfr_nbvar = 20 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("LAMBDA_cl33", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6080:/ getarg");
 /*_.CURFLDA__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_LONG)
    goto meltlab_endgetargs;
 /*_#CURK__L1*/ meltfnum[0] = meltxargtab_[0].meltbp_long;
  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6081:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L2*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CURFLDA__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_FIELDASSIGN */
					    meltfrout->tabval[0])));;
      MELT_LOCATION ("warmelt-normal.melt:6081:/ cond");
      /*cond */ if ( /*_#IS_A__L2*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V4*/ meltfptr[3] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6081:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check curflda"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6081) ? (6081) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V4*/ meltfptr[3] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V3*/ meltfptr[2] = /*_._IFELSE___V4*/ meltfptr[3];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6081:/ clear");
	     /*clear *//*_#IS_A__L2*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V4*/ meltfptr[3] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V3*/ meltfptr[2] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6082:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CURFLDA__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.CURFLOC__V6*/ meltfptr[5] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6083:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CURFLDA__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SFLA_FIELD");
  /*_.CURFIELD__V7*/ meltfptr[6] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6084:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CURFLDA__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 3, "SFLA_EXPR");
  /*_.CUREXP__V8*/ meltfptr[7] = slot;
    };
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#NULL__L3*/ meltfnum[1] =
      (( /*_.CURFLOC__V6*/ meltfptr[5]) == NULL);;
    MELT_LOCATION ("warmelt-normal.melt:6085:/ cond");
    /*cond */ if ( /*_#NULL__L3*/ meltfnum[1])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  /*^compute */
	  /*_.CURFLOC__V6*/ meltfptr[5] = /*_.SETQ___V10*/ meltfptr[9] =
	    ( /*~SLOC */ meltfclos->tabval[0]);;
	  /*_._IF___V9*/ meltfptr[8] = /*_.SETQ___V10*/ meltfptr[9];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6085:/ clear");
	     /*clear *//*_.SETQ___V10*/ meltfptr[9] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

  /*_._IF___V9*/ meltfptr[8] = NULL;;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6086:/ quasiblock");


    /*^multimsend */
    /*multimsend */
    {
      union meltparam_un argtab[3];
      union meltparam_un restab[1];
      memset (&argtab, 0, sizeof (argtab));
      memset (&restab, 0, sizeof (restab));
      /*^multimsend.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & ( /*~ENV */ meltfclos->tabval[1]);	/*^multimsend.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & ( /*~NCX */ meltfclos->tabval[2]);	/*^multimsend.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.CURFLOC__V6*/ meltfptr[5];
      /*^multimsend.xres */
      restab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBIND__V13*/ meltfptr[12];	/*^multimsend.send */
      /*_.NEXP__V12*/ meltfptr[11] =
	meltgc_send ((melt_ptr_t) ( /*_.CUREXP__V8*/ meltfptr[7]),
		     ((melt_ptr_t)
		      (( /*!NORMAL_EXP */ meltfrout->tabval[1]))),
		     (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		     argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6089:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[3] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6089:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[3])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[4] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6089:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[4];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6089;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normletrec_instance nexp";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NEXP__V12*/ meltfptr[11];
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[2])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V14*/ meltfptr[13] =
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6089:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[4] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V14*/ meltfptr[13] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6089:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[3] = 0;
      /*^clear */
	     /*clear *//*_._IF___V14*/ meltfptr[13] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6090:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L6*/ meltfnum[4] =
	(( /*_.NBIND__V13*/ meltfptr[12]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBIND__V13*/ meltfptr[12])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:6090:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L6*/ meltfnum[4])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V17*/ meltfptr[13] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6090:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nbind"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6090) ? (6090) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V17*/ meltfptr[13] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V16*/ meltfptr[14] = /*_._IFELSE___V17*/ meltfptr[13];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6090:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L6*/ meltfnum[4] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V17*/ meltfptr[13] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V16*/ meltfptr[14] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6091:/ apply");
    /*apply */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^apply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NBIND__V13*/ meltfptr[12];
      /*_.LIST_APPEND2LIST__V18*/ meltfptr[13] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!LIST_APPEND2LIST */ meltfrout->tabval[3])),
		    (melt_ptr_t) (( /*~BINDLIST */ meltfclos->tabval[3])),
		    (MELTBPARSTR_PTR ""), argtab, "",
		    (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6092:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NREP_FIELDASSIGN */
					     meltfrout->tabval[4])), (3),
			      "CLASS_NREP_FIELDASSIGN");
  /*_.INST__V20*/ meltfptr[19] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (0),
			  ( /*_.CURFLOC__V6*/ meltfptr[5]), "NREP_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NFLA_FIELD",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (1),
			  ( /*_.CURFIELD__V7*/ meltfptr[6]), "NFLA_FIELD");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NFLA_VAL",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V20*/ meltfptr[19])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (2),
			  ( /*_.NEXP__V12*/ meltfptr[11]), "NFLA_VAL");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V20*/ meltfptr[19],
				  "newly made instance");
    ;
    /*_.INST___V19*/ meltfptr[18] = /*_.INST__V20*/ meltfptr[19];;

    MELT_CHECK_SIGNAL ();
    ;
    /*^compute */
    /*_.MULTI___V11*/ meltfptr[9] = /*_.INST___V19*/ meltfptr[18];;

    MELT_LOCATION ("warmelt-normal.melt:6086:/ clear");
	   /*clear *//*_.IFCPP___V16*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.LIST_APPEND2LIST__V18*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_.INST___V19*/ meltfptr[18] = 0;

    /*^clear */
	   /*clear *//*_.NBIND__V13*/ meltfptr[12] = 0;
    /*_.LET___V5*/ meltfptr[3] = /*_.MULTI___V11*/ meltfptr[9];;

    MELT_LOCATION ("warmelt-normal.melt:6082:/ clear");
	   /*clear *//*_.CURFLOC__V6*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.CURFIELD__V7*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.CUREXP__V8*/ meltfptr[7] = 0;
    /*^clear */
	   /*clear *//*_#NULL__L3*/ meltfnum[1] = 0;
    /*^clear */
	   /*clear *//*_._IF___V9*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.MULTI___V11*/ meltfptr[9] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6080:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V5*/ meltfptr[3];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6080:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V3*/ meltfptr[2] = 0;
    /*^clear */
	   /*clear *//*_.LET___V5*/ meltfptr[3] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("LAMBDA_cl33", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_113_WARMELTmiNORMAL_LAMBDA_cl33 */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *
					     meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un *
					     meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 153
    melt_ptr_t mcfr_varptr[153];
#define MELTFRAM_NBVARNUM 72
    long mcfr_varnum[72];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC_st *)
	meltfirstargp_;
      /* use arguments meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 153; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC nbval 153*/
  meltfram__.mcfr_nbvar = 153 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMEXP_LETREC", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6127:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.PSLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.PSLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6128:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L1*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_LETREC */
					    meltfrout->tabval[0])));;
      MELT_LOCATION ("warmelt-normal.melt:6128:/ cond");
      /*cond */ if ( /*_#IS_A__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V7*/ meltfptr[6] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6128:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check letrec recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6128) ? (6128) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V7*/ meltfptr[6] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V6*/ meltfptr[5] = /*_._IFELSE___V7*/ meltfptr[6];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6128:/ clear");
	     /*clear *//*_#IS_A__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V7*/ meltfptr[6] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V6*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6129:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L2*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6129:/ cond");
      /*cond */ if ( /*_#IS_A__L2*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6129:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check env"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6129) ? (6129) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6129:/ clear");
	     /*clear *//*_#IS_A__L2*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6130:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:6130:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[10] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6130:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6130) ? (6130) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[8] = /*_._IFELSE___V11*/ meltfptr[10];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6130:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[8] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6131:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6131:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[4] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6131:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[4];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6131;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec recv";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V13*/ meltfptr[12] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V12*/ meltfptr[10] =
	      /*_.MELT_DEBUG_FUN__V13*/ meltfptr[12];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6131:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[4] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V13*/ meltfptr[12] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V12*/ meltfptr[10] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6131:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V12*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6132:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.SLOC__V15*/ meltfptr[10] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6133:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SLET_BINDINGS");
  /*_.SBINDS__V16*/ meltfptr[15] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6134:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 3, "SLET_BODY");
  /*_.SBODY__V17*/ meltfptr[16] = slot;
    };
    ;
 /*_#NBBIND__L6*/ meltfnum[4] =
      (melt_multiple_length
       ((melt_ptr_t) ( /*_.SBINDS__V16*/ meltfptr[15])));;
    /*^compute */
 /*_.BINDLIST__V18*/ meltfptr[17] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[4]))));;
    /*^compute */
 /*_.CONSBINDLIST__V19*/ meltfptr[18] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[4]))));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6138:/ apply");
    /*apply */
    {
      /*_.NEWENV__V20*/ meltfptr[19] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!FRESH_ENV */ meltfrout->tabval[5])),
		    (melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;
 /*_#stI__L7*/ meltfnum[0] =
      (( /*_#NBBIND__L6*/ meltfnum[4]) * (2));;
    /*^compute */
 /*_#MAPLEN__L8*/ meltfnum[7] =
      ((5) + ( /*_#stI__L7*/ meltfnum[0]));;
    /*^compute */
 /*_.SYMBEXPRMAP__V21*/ meltfptr[20] =
      (meltgc_new_mapobjects
       ((meltobject_ptr_t) (( /*!DISCR_MAP_OBJECTS */ meltfrout->tabval[6])),
	( /*_#MAPLEN__L8*/ meltfnum[7])));;
    /*^compute */
 /*_.SYMBINDMAP__V22*/ meltfptr[21] =
      (meltgc_new_mapobjects
       ((meltobject_ptr_t) (( /*!DISCR_MAP_OBJECTS */ meltfrout->tabval[6])),
	( /*_#MAPLEN__L8*/ meltfnum[7])));;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit1__EACHTUP */
      long meltcit1__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.SBINDS__V16*/ meltfptr[15]);
      for ( /*_#SBIX__L9*/ meltfnum[8] = 0;
	   ( /*_#SBIX__L9*/ meltfnum[8] >= 0)
	   && ( /*_#SBIX__L9*/ meltfnum[8] < meltcit1__EACHTUP_ln);
	/*_#SBIX__L9*/ meltfnum[8]++)
	{
	  /*_.CURSBIND__V23*/ meltfptr[22] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.SBINDS__V16*/ meltfptr[15]),
			       /*_#SBIX__L9*/ meltfnum[8]);




#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6147:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L10*/ meltfnum[9] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6147:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L11*/ meltfnum[10] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6147:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L11*/ meltfnum[10];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6147;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normexp_letrec cursbind";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURSBIND__V23*/ meltfptr[22];
		    /*_.MELT_DEBUG_FUN__V25*/ meltfptr[24] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V24*/ meltfptr[23] =
		    /*_.MELT_DEBUG_FUN__V25*/ meltfptr[24];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6147:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[10] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V25*/ meltfptr[24] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V24*/ meltfptr[23] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6147:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[9] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V24*/ meltfptr[23] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6148:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L12*/ meltfnum[10] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURSBIND__V23*/ meltfptr[22]),
				   (melt_ptr_t) (( /*!CLASS_SOURCE_LETREC_BINDING */ meltfrout->tabval[7])));;
	    MELT_LOCATION ("warmelt-normal.melt:6148:/ cond");
	    /*cond */ if ( /*_#IS_A__L12*/ meltfnum[10])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V27*/ meltfptr[23] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6148:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check cursbind"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6148) ? (6148) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V27*/ meltfptr[23] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V26*/ meltfptr[24] =
	      /*_._IFELSE___V27*/ meltfptr[23];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6148:/ clear");
	      /*clear *//*_#IS_A__L12*/ meltfnum[10] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V27*/ meltfptr[23] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V26*/ meltfptr[24] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6149:/ quasiblock");


	  /*^getslot */
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V23*/ meltfptr[22]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOCA_LOCATION__V29*/ meltfptr[28] = slot;
	  };
	  ;
	  /*^cond */
	  /*cond */ if ( /*_.LOCA_LOCATION__V29*/ meltfptr[28])	/*then */
	    {
	      /*^cond.then */
	      /*_.LOCB__V30*/ meltfptr[29] =
		/*_.LOCA_LOCATION__V29*/ meltfptr[28];;
	    }
	  else
	    {
	      MELT_LOCATION ("warmelt-normal.melt:6149:/ cond.else");

	      /*^block */
	      /*anyblock */
	      {

		/*^cond */
		/*cond */ if ( /*_.SLOC__V15*/ meltfptr[10])	/*then */
		  {
		    /*^cond.then */
		    /*_.OR___V31*/ meltfptr[30] =
		      /*_.SLOC__V15*/ meltfptr[10];;
		  }
		else
		  {
		    MELT_LOCATION ("warmelt-normal.melt:6149:/ cond.else");

		    /*_.OR___V31*/ meltfptr[30] =
		      /*_.PSLOC__V5*/ meltfptr[4];;
		  }
		;
		/*^compute */
		/*_.LOCB__V30*/ meltfptr[29] = /*_.OR___V31*/ meltfptr[30];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6149:/ clear");
	      /*clear *//*_.OR___V31*/ meltfptr[30] = 0;
	      }
	      ;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6150:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V23*/ meltfptr[22]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 3, "SLETB_BINDER");
   /*_.CURSYMB__V32*/ meltfptr[30] = slot;
	  };
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6151:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V23*/ meltfptr[22]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 4, "SLETB_EXPR");
   /*_.CUREXPR__V33*/ meltfptr[32] = slot;
	  };
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6153:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L13*/ meltfnum[9] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6153:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L13*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L14*/ meltfnum[10] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6153:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[7];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L14*/ meltfnum[10];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6153;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normexp_letrec curexpr=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CUREXPR__V33*/ meltfptr[32];
		    /*^apply.arg */
		    argtab[5].meltbp_cstring = " cursymb=";
		    /*^apply.arg */
		    argtab[6].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURSYMB__V32*/ meltfptr[30];
		    /*_.MELT_DEBUG_FUN__V35*/ meltfptr[34] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V34*/ meltfptr[33] =
		    /*_.MELT_DEBUG_FUN__V35*/ meltfptr[34];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6153:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L14*/ meltfnum[10] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V35*/ meltfptr[34] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V34*/ meltfptr[33] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6153:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L13*/ meltfnum[9] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V34*/ meltfptr[33] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6154:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L15*/ meltfnum[10] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CUREXPR__V33*/ meltfptr[32]),
				   (melt_ptr_t) (( /*!CLASS_SOURCE */
						  meltfrout->tabval[8])));;
	    MELT_LOCATION ("warmelt-normal.melt:6154:/ cond");
	    /*cond */ if ( /*_#IS_A__L15*/ meltfnum[10])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V37*/ meltfptr[33] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6154:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curexpr"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6154) ? (6154) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V37*/ meltfptr[33] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V36*/ meltfptr[34] =
	      /*_._IFELSE___V37*/ meltfptr[33];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6154:/ clear");
	      /*clear *//*_#IS_A__L15*/ meltfnum[10] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V37*/ meltfptr[33] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V36*/ meltfptr[34] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6155:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L16*/ meltfnum[9] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURSYMB__V32*/ meltfptr[30]),
				   (melt_ptr_t) (( /*!CLASS_SYMBOL */
						  meltfrout->tabval[9])));;
	    MELT_LOCATION ("warmelt-normal.melt:6155:/ cond");
	    /*cond */ if ( /*_#IS_A__L16*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V39*/ meltfptr[38] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6155:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check cursymb"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6155) ? (6155) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V39*/ meltfptr[38] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V38*/ meltfptr[33] =
	      /*_._IFELSE___V39*/ meltfptr[38];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6155:/ clear");
	      /*clear *//*_#IS_A__L16*/ meltfnum[9] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V39*/ meltfptr[38] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V38*/ meltfptr[33] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6156:/ locexp");
	    meltgc_put_mapobjects ((meltmapobjects_ptr_t)
				   ( /*_.SYMBEXPRMAP__V21*/ meltfptr[20]),
				   (meltobject_ptr_t) ( /*_.CURSYMB__V32*/
						       meltfptr[30]),
				   (melt_ptr_t) ( /*_.CUREXPR__V33*/
						 meltfptr[32]));
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6157:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^msend */
	  /*msend */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.CURSYMB__V32*/ meltfptr[30];
	    /*^ojbmsend.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
	    /*^ojbmsend.arg */
	    argtab[2].meltbp_aptr =
	      (melt_ptr_t *) & /*_.LOCB__V30*/ meltfptr[29];
	    /*_.CURCONSBIND__V41*/ meltfptr[40] =
	      meltgc_send ((melt_ptr_t) ( /*_.CUREXPR__V33*/ meltfptr[32]),
			   (melt_ptr_t) (( /*!PREPARE_CONSTRUCTOR_BINDING */
					  meltfrout->tabval[10])),
			   (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
			    ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6159:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L17*/ meltfnum[10] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6159:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L17*/ meltfnum[10])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L18*/ meltfnum[9] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6159:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L18*/ meltfnum[9];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6159;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normexp_letrec curconsbind";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURCONSBIND__V41*/ meltfptr[40];
		    /*_.MELT_DEBUG_FUN__V43*/ meltfptr[42] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V42*/ meltfptr[41] =
		    /*_.MELT_DEBUG_FUN__V43*/ meltfptr[42];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6159:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L18*/ meltfnum[9] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V43*/ meltfptr[42] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V42*/ meltfptr[41] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6159:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L17*/ meltfnum[10] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V42*/ meltfptr[41] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
  /*_#IS_A__L19*/ meltfnum[9] =
	    melt_is_instance_of ((melt_ptr_t)
				 ( /*_.CURCONSBIND__V41*/ meltfptr[40]),
				 (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])));;
	  MELT_LOCATION ("warmelt-normal.melt:6160:/ cond");
	  /*cond */ if ( /*_#IS_A__L19*/ meltfnum[9])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {


#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6161:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#MELT_NEED_DBG__L20*/ meltfnum[10] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:6161:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L20*/ meltfnum[10])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	/*_#MELT_CALLCOUNT__L21*/ meltfnum[20] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:6161:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[5];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L21*/ meltfnum[20];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 6161;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "normexp_letrec curconsbind plain consbind";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CURCONSBIND__V41*/
			    meltfptr[40];
			  /*_.MELT_DEBUG_FUN__V46*/ meltfptr[45] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[3])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V45*/ meltfptr[41] =
			  /*_.MELT_DEBUG_FUN__V46*/ meltfptr[45];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6161:/ clear");
		  /*clear *//*_#MELT_CALLCOUNT__L21*/ meltfnum[20] = 0;
			/*^clear */
		  /*clear *//*_.MELT_DEBUG_FUN__V46*/ meltfptr[45] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

       /*_._IF___V45*/ meltfptr[41] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:6161:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		/*clear *//*_#MELT_NEED_DBG__L20*/ meltfnum[10] = 0;
		  /*^clear */
		/*clear *//*_._IF___V45*/ meltfptr[41] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

		{
		  MELT_LOCATION ("warmelt-normal.melt:6162:/ locexp");
		  meltgc_append_list ((melt_ptr_t)
				      ( /*_.CONSBINDLIST__V19*/ meltfptr[18]),
				      (melt_ptr_t) ( /*_.CURCONSBIND__V41*/
						    meltfptr[40]));
		}
		;

		{
		  MELT_LOCATION ("warmelt-normal.melt:6163:/ locexp");
		  meltgc_put_mapobjects ((meltmapobjects_ptr_t)
					 ( /*_.SYMBINDMAP__V22*/
					  meltfptr[21]),
					 (meltobject_ptr_t) ( /*_.CURSYMB__V32*/ meltfptr[30]),
					 (melt_ptr_t) ( /*_.CURCONSBIND__V41*/
						       meltfptr[40]));
		}
		;
		MELT_LOCATION ("warmelt-normal.melt:6160:/ quasiblock");


		/*epilog */
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

	      /*^block */
	      /*anyblock */
	      {


		MELT_CHECK_SIGNAL ();
		;
    /*_#IS_MULTIPLE__L22*/ meltfnum[20] =
		  (melt_magic_discr
		   ((melt_ptr_t) ( /*_.CURCONSBIND__V41*/ meltfptr[40])) ==
		   MELTOBMAG_MULTIPLE);;
		MELT_LOCATION ("warmelt-normal.melt:6165:/ cond");
		/*cond */ if ( /*_#IS_MULTIPLE__L22*/ meltfnum[20])	/*then */
		  {
		    /*^cond.then */
		    /*^block */
		    /*anyblock */
		    {


#if MELT_HAVE_DEBUG
		      MELT_LOCATION ("warmelt-normal.melt:6166:/ cppif.then");
		      /*^block */
		      /*anyblock */
		      {


			{
			  /*^locexp */
			  /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
			  melt_dbgcounter++;
#endif
			  ;
			}
			;

			MELT_CHECK_SIGNAL ();
			;
	/*_#MELT_NEED_DBG__L23*/ meltfnum[10] =
			  /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
			  ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
			  0	/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
			  ;;
			MELT_LOCATION ("warmelt-normal.melt:6166:/ cond");
			/*cond */ if ( /*_#MELT_NEED_DBG__L23*/ meltfnum[10])	/*then */
			  {
			    /*^cond.then */
			    /*^block */
			    /*anyblock */
			    {

	  /*_#MELT_CALLCOUNT__L24*/ meltfnum[23] =
				/* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
				meltcallcount	/* melt_callcount debugging */
#else
				0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
				;;

			      MELT_CHECK_SIGNAL ();
			      ;
			      MELT_LOCATION
				("warmelt-normal.melt:6166:/ apply");
			      /*apply */
			      {
				union meltparam_un argtab[5];
				memset (&argtab, 0, sizeof (argtab));
				/*^apply.arg */
				argtab[0].meltbp_long =
				  /*_#MELT_CALLCOUNT__L24*/ meltfnum[23];
				/*^apply.arg */
				argtab[1].meltbp_cstring =
				  "warmelt-normal.melt";
				/*^apply.arg */
				argtab[2].meltbp_long = 6166;
				/*^apply.arg */
				argtab[3].meltbp_cstring =
				  "normexp_letrec curconsbind multiple";
				/*^apply.arg */
				argtab[4].meltbp_aptr =
				  (melt_ptr_t *) & /*_.CURCONSBIND__V41*/
				  meltfptr[40];
				/*_.MELT_DEBUG_FUN__V49*/ meltfptr[48] =
				  melt_apply ((meltclosure_ptr_t)
					      (( /*!MELT_DEBUG_FUN */
						meltfrout->tabval[3])),
					      (melt_ptr_t) (( /*nil */ NULL)),
					      (MELTBPARSTR_LONG
					       MELTBPARSTR_CSTRING
					       MELTBPARSTR_LONG
					       MELTBPARSTR_CSTRING
					       MELTBPARSTR_PTR ""), argtab,
					      "", (union meltparam_un *) 0);
			      }
			      ;
			      /*_._IF___V48*/ meltfptr[41] =
				/*_.MELT_DEBUG_FUN__V49*/ meltfptr[48];;
			      /*epilog */

			      MELT_LOCATION
				("warmelt-normal.melt:6166:/ clear");
		    /*clear *//*_#MELT_CALLCOUNT__L24*/ meltfnum[23]
				= 0;
			      /*^clear */
		    /*clear *//*_.MELT_DEBUG_FUN__V49*/ meltfptr[48]
				= 0;
			    }
			    ;
			  }
			else
			  {	/*^cond.else */

	 /*_._IF___V48*/ meltfptr[41] = NULL;;
			  }
			;

			{
			  MELT_LOCATION ("warmelt-normal.melt:6166:/ locexp");
			  /*void */ (void) 0;
			}
			;
			/*^quasiblock */


			/*epilog */

			/*^clear */
		  /*clear *//*_#MELT_NEED_DBG__L23*/ meltfnum[10] = 0;
			/*^clear */
		  /*clear *//*_._IF___V48*/ meltfptr[41] = 0;
		      }

#else /*MELT_HAVE_DEBUG */
		      /*^cppif.else */
		      /*^block */
		      /*anyblock */
		      {


			{
			  /*^locexp */
			  /*void */ (void) 0;
			}
			;
			/*epilog */
		      }

#endif /*MELT_HAVE_DEBUG */
		      ;
		      /*citerblock FOREACH_IN_MULTIPLE */
		      {
			/* start foreach_in_multiple meltcit2__EACHTUP */
			long meltcit2__EACHTUP_ln =
			  melt_multiple_length ((melt_ptr_t)
						/*_.CURCONSBIND__V41*/
						meltfptr[40]);
			for ( /*_#SUBIX__L25*/ meltfnum[23] = 0;
			     ( /*_#SUBIX__L25*/ meltfnum[23] >= 0)
			     && ( /*_#SUBIX__L25*/ meltfnum[23] <
				 meltcit2__EACHTUP_ln);
	/*_#SUBIX__L25*/ meltfnum[23]++)
			  {
			    /*_.SUBCONSBIND__V50*/ meltfptr[48] =
			      melt_multiple_nth ((melt_ptr_t)
						 ( /*_.CURCONSBIND__V41*/
						  meltfptr[40]),
						 /*_#SUBIX__L25*/
						 meltfnum[23]);




#if MELT_HAVE_DEBUG
			    MELT_LOCATION
			      ("warmelt-normal.melt:6170:/ cppif.then");
			    /*^block */
			    /*anyblock */
			    {


			      MELT_CHECK_SIGNAL ();
			      ;
	 /*_#IS_A__L26*/ meltfnum[10] =
				melt_is_instance_of ((melt_ptr_t)
						     ( /*_.SUBCONSBIND__V50*/
						      meltfptr[48]),
						     (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])));;
			      MELT_LOCATION
				("warmelt-normal.melt:6170:/ cond");
			      /*cond */ if ( /*_#IS_A__L26*/ meltfnum[10])	/*then */
				{
				  /*^cond.then */
				  /*_._IFELSE___V52*/ meltfptr[51] =
				    ( /*nil */ NULL);;
				}
			      else
				{
				  MELT_LOCATION
				    ("warmelt-normal.melt:6170:/ cond.else");

				  /*^block */
				  /*anyblock */
				  {




				    {
				      /*^locexp */
				      melt_assert_failed (("normexp_letrec check subconsbind"), ("warmelt-normal.melt") ? ("warmelt-normal.melt") : __FILE__, (6170) ? (6170) : __LINE__, __FUNCTION__);
				      ;
				    }
				    ;
		     /*clear *//*_._IFELSE___V52*/ meltfptr[51]
				      = 0;
				    /*epilog */
				  }
				  ;
				}
			      ;
			      /*^compute */
			      /*_.IFCPP___V51*/ meltfptr[41] =
				/*_._IFELSE___V52*/ meltfptr[51];;
			      /*epilog */

			      MELT_LOCATION
				("warmelt-normal.melt:6170:/ clear");
		   /*clear *//*_#IS_A__L26*/ meltfnum[10] = 0;
			      /*^clear */
		   /*clear *//*_._IFELSE___V52*/ meltfptr[51] = 0;
			    }

#else /*MELT_HAVE_DEBUG */
			    /*^cppif.else */
			    /*_.IFCPP___V51*/ meltfptr[41] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
			    ;

			    {
			      MELT_LOCATION
				("warmelt-normal.melt:6172:/ locexp");
			      meltgc_append_list ((melt_ptr_t)
						  ( /*_.CONSBINDLIST__V19*/
						   meltfptr[18]),
						  (melt_ptr_t) ( /*_.SUBCONSBIND__V50*/ meltfptr[48]));
			    }
			    ;
			    if ( /*_#SUBIX__L25*/ meltfnum[23] < 0)
			      break;
			  }	/* end  foreach_in_multiple meltcit2__EACHTUP */

			/*citerepilog */

			MELT_LOCATION ("warmelt-normal.melt:6167:/ clear");
		 /*clear *//*_.SUBCONSBIND__V50*/ meltfptr[48] = 0;
			/*^clear */
		 /*clear *//*_#SUBIX__L25*/ meltfnum[23] = 0;
			/*^clear */
		 /*clear *//*_.IFCPP___V51*/ meltfptr[41] = 0;
		      }		/*endciterblock FOREACH_IN_MULTIPLE */
		      ;

		      {
			MELT_LOCATION ("warmelt-normal.melt:6173:/ locexp");
			meltgc_put_mapobjects ((meltmapobjects_ptr_t)
					       ( /*_.SYMBINDMAP__V22*/
						meltfptr[21]),
					       (meltobject_ptr_t) ( /*_.CURSYMB__V32*/ meltfptr[30]),
					       (melt_ptr_t) ( /*_.CURCONSBIND__V41*/ meltfptr[40]));
		      }
		      ;
		      MELT_LOCATION ("warmelt-normal.melt:6165:/ quasiblock");


		      /*epilog */
		    }
		    ;
		  }
		else
		  {		/*^cond.else */

		    /*^block */
		    /*anyblock */
		    {


		      MELT_CHECK_SIGNAL ();
		      ;
      /*_#IS_LIST__L27*/ meltfnum[10] =
			(melt_magic_discr
			 ((melt_ptr_t) ( /*_.CURCONSBIND__V41*/ meltfptr[40]))
			 == MELTOBMAG_LIST);;
		      MELT_LOCATION ("warmelt-normal.melt:6175:/ cond");
		      /*cond */ if ( /*_#IS_LIST__L27*/ meltfnum[10])	/*then */
			{
			  /*^cond.then */
			  /*^block */
			  /*anyblock */
			  {


#if MELT_HAVE_DEBUG
			    MELT_LOCATION
			      ("warmelt-normal.melt:6176:/ cppif.then");
			    /*^block */
			    /*anyblock */
			    {


			      {
				/*^locexp */
				/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
				melt_dbgcounter++;
#endif
				;
			      }
			      ;

			      MELT_CHECK_SIGNAL ();
			      ;
	  /*_#MELT_NEED_DBG__L28*/ meltfnum[27] =
				/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
				( /*melt_need_dbg */
				 melt_need_debug ((int) 0))
#else
				0	/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
				;;
			      MELT_LOCATION
				("warmelt-normal.melt:6176:/ cond");
			      /*cond */ if ( /*_#MELT_NEED_DBG__L28*/ meltfnum[27])	/*then */
				{
				  /*^cond.then */
				  /*^block */
				  /*anyblock */
				  {

	    /*_#MELT_CALLCOUNT__L29*/ meltfnum[28] =
				      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
				      meltcallcount	/* melt_callcount debugging */
#else
				      0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
				      ;;

				    MELT_CHECK_SIGNAL ();
				    ;
				    MELT_LOCATION
				      ("warmelt-normal.melt:6176:/ apply");
				    /*apply */
				    {
				      union meltparam_un argtab[5];
				      memset (&argtab, 0, sizeof (argtab));
				      /*^apply.arg */
				      argtab[0].meltbp_long =
					/*_#MELT_CALLCOUNT__L29*/
					meltfnum[28];
				      /*^apply.arg */
				      argtab[1].meltbp_cstring =
					"warmelt-normal.melt";
				      /*^apply.arg */
				      argtab[2].meltbp_long = 6176;
				      /*^apply.arg */
				      argtab[3].meltbp_cstring =
					"normexp_letrec curconsbind list";
				      /*^apply.arg */
				      argtab[4].meltbp_aptr =
					(melt_ptr_t *) &
					/*_.CURCONSBIND__V41*/ meltfptr[40];
				      /*_.MELT_DEBUG_FUN__V55*/ meltfptr[54] =
					melt_apply ((meltclosure_ptr_t)
						    (( /*!MELT_DEBUG_FUN */
						      meltfrout->tabval[3])),
						    (melt_ptr_t) (( /*nil */
								   NULL)),
						    (MELTBPARSTR_LONG
						     MELTBPARSTR_CSTRING
						     MELTBPARSTR_LONG
						     MELTBPARSTR_CSTRING
						     MELTBPARSTR_PTR ""),
						    argtab, "",
						    (union meltparam_un *) 0);
				    }
				    ;
				    /*_._IF___V54*/ meltfptr[53] =
				      /*_.MELT_DEBUG_FUN__V55*/ meltfptr[54];;
				    /*epilog */

				    MELT_LOCATION
				      ("warmelt-normal.melt:6176:/ clear");
		      /*clear *//*_#MELT_CALLCOUNT__L29*/
				      meltfnum[28] = 0;
				    /*^clear */
		      /*clear *//*_.MELT_DEBUG_FUN__V55*/
				      meltfptr[54] = 0;
				  }
				  ;
				}
			      else
				{	/*^cond.else */

	   /*_._IF___V54*/ meltfptr[53] = NULL;;
				}
			      ;

			      {
				MELT_LOCATION
				  ("warmelt-normal.melt:6176:/ locexp");
				/*void */ (void) 0;
			      }
			      ;
			      /*^quasiblock */


			      /*epilog */

			      /*^clear */
		    /*clear *//*_#MELT_NEED_DBG__L28*/ meltfnum[27]
				= 0;
			      /*^clear */
		    /*clear *//*_._IF___V54*/ meltfptr[53] = 0;
			    }

#else /*MELT_HAVE_DEBUG */
			    /*^cppif.else */
			    /*^block */
			    /*anyblock */
			    {


			      {
				/*^locexp */
				/*void */ (void) 0;
			      }
			      ;
			      /*epilog */
			    }

#endif /*MELT_HAVE_DEBUG */
			    ;
			    /*citerblock FOREACH_IN_LIST */
			    {
			      /* start foreach_in_list meltcit3__EACHLIST */
			      for ( /*_.SUBCONSPAIR__V56*/ meltfptr[54] =
				   melt_list_first ((melt_ptr_t)
						    /*_.CURCONSBIND__V41*/
						    meltfptr[40]);
				   melt_magic_discr ((melt_ptr_t)
						     /*_.SUBCONSPAIR__V56*/
						     meltfptr[54]) ==
				   MELTOBMAG_PAIR;
				   /*_.SUBCONSPAIR__V56*/ meltfptr[54] =
				   melt_pair_tail ((melt_ptr_t)
						   /*_.SUBCONSPAIR__V56*/
						   meltfptr[54]))
				{
				  /*_.SUBCONSBIND__V57*/ meltfptr[53] =
				    melt_pair_head ((melt_ptr_t)
						    /*_.SUBCONSPAIR__V56*/
						    meltfptr[54]);



#if MELT_HAVE_DEBUG
				  MELT_LOCATION
				    ("warmelt-normal.melt:6180:/ cppif.then");
				  /*^block */
				  /*anyblock */
				  {


				    MELT_CHECK_SIGNAL ();
				    ;
	   /*_#IS_A__L30*/ meltfnum[28] =
				      melt_is_instance_of ((melt_ptr_t)
							   ( /*_.SUBCONSBIND__V57*/ meltfptr[53]), (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])));;
				    MELT_LOCATION
				      ("warmelt-normal.melt:6180:/ cond");
				    /*cond */ if ( /*_#IS_A__L30*/ meltfnum[28])	/*then */
				      {
					/*^cond.then */
					/*_._IFELSE___V59*/ meltfptr[58] =
					  ( /*nil */ NULL);;
				      }
				    else
				      {
					MELT_LOCATION
					  ("warmelt-normal.melt:6180:/ cond.else");

					/*^block */
					/*anyblock */
					{




					  {
					    /*^locexp */
					    melt_assert_failed (("normexp_letrec check subconsbind"), ("warmelt-normal.melt") ? ("warmelt-normal.melt") : __FILE__, (6180) ? (6180) : __LINE__, __FUNCTION__);
					    ;
					  }
					  ;
		       /*clear *//*_._IFELSE___V59*/
					    meltfptr[58] = 0;
					  /*epilog */
					}
					;
				      }
				    ;
				    /*^compute */
				    /*_.IFCPP___V58*/ meltfptr[57] =
				      /*_._IFELSE___V59*/ meltfptr[58];;
				    /*epilog */

				    MELT_LOCATION
				      ("warmelt-normal.melt:6180:/ clear");
		     /*clear *//*_#IS_A__L30*/ meltfnum[28] =
				      0;
				    /*^clear */
		     /*clear *//*_._IFELSE___V59*/ meltfptr[58]
				      = 0;
				  }

#else /*MELT_HAVE_DEBUG */
				  /*^cppif.else */
				  /*_.IFCPP___V58*/ meltfptr[57] =
				    ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
				  ;

				  {
				    MELT_LOCATION
				      ("warmelt-normal.melt:6182:/ locexp");
				    meltgc_append_list ((melt_ptr_t)
							( /*_.CONSBINDLIST__V19*/ meltfptr[18]), (melt_ptr_t) ( /*_.SUBCONSBIND__V57*/ meltfptr[53]));
				  }
				  ;
				}	/* end foreach_in_list meltcit3__EACHLIST */
     /*_.SUBCONSPAIR__V56*/ meltfptr[54] = NULL;
     /*_.SUBCONSBIND__V57*/ meltfptr[53] = NULL;


			      /*citerepilog */

			      MELT_LOCATION
				("warmelt-normal.melt:6177:/ clear");
		   /*clear *//*_.SUBCONSPAIR__V56*/ meltfptr[54] =
				0;
			      /*^clear */
		   /*clear *//*_.SUBCONSBIND__V57*/ meltfptr[53] =
				0;
			      /*^clear */
		   /*clear *//*_.IFCPP___V58*/ meltfptr[57] = 0;
			    }	/*endciterblock FOREACH_IN_LIST */
			    ;

			    MELT_CHECK_SIGNAL ();
			    ;
			    MELT_LOCATION
			      ("warmelt-normal.melt:6183:/ apply");
			    /*apply */
			    {
			      union meltparam_un argtab[1];
			      memset (&argtab, 0, sizeof (argtab));
			      /*^apply.arg */
			      argtab[0].meltbp_aptr =
				(melt_ptr_t *) & ( /*!DISCR_MULTIPLE */
						  meltfrout->tabval[13]);
			      /*_.LIST_TO_MULTIPLE__V60*/ meltfptr[58] =
				melt_apply ((meltclosure_ptr_t)
					    (( /*!LIST_TO_MULTIPLE */
					      meltfrout->tabval[12])),
					    (melt_ptr_t) ( /*_.CURCONSBIND__V41*/ meltfptr[40]), (MELTBPARSTR_PTR ""), argtab, "", (union meltparam_un *) 0);
			    }
			    ;

			    {
			      /*^locexp */
			      meltgc_put_mapobjects ((meltmapobjects_ptr_t)
						     ( /*_.SYMBINDMAP__V22*/
						      meltfptr[21]),
						     (meltobject_ptr_t) ( /*_.CURSYMB__V32*/ meltfptr[30]),
						     (melt_ptr_t) ( /*_.LIST_TO_MULTIPLE__V60*/ meltfptr[58]));
			    }
			    ;
			    MELT_LOCATION
			      ("warmelt-normal.melt:6175:/ quasiblock");


			    /*epilog */

			    /*^clear */
		  /*clear *//*_.LIST_TO_MULTIPLE__V60*/ meltfptr[58]
			      = 0;
			  }
			  ;
			}
		      else
			{	/*^cond.else */

			  /*^block */
			  /*anyblock */
			  {


#if MELT_HAVE_DEBUG
			    MELT_LOCATION
			      ("warmelt-normal.melt:6186:/ cppif.then");
			    /*^block */
			    /*anyblock */
			    {


			      MELT_CHECK_SIGNAL ();
			      ;
			      /*^cond */
			      /*cond */ if (( /*nil */ NULL))	/*then */
				{
				  /*^cond.then */
				  /*_._IFELSE___V62*/ meltfptr[61] =
				    ( /*nil */ NULL);;
				}
			      else
				{
				  MELT_LOCATION
				    ("warmelt-normal.melt:6186:/ cond.else");

				  /*^block */
				  /*anyblock */
				  {




				    {
				      /*^locexp */
				      melt_assert_failed (("normexp_letrec bad curconsbind"), ("warmelt-normal.melt") ? ("warmelt-normal.melt") : __FILE__, (6186) ? (6186) : __LINE__, __FUNCTION__);
				      ;
				    }
				    ;
		      /*clear *//*_._IFELSE___V62*/ meltfptr[61]
				      = 0;
				    /*epilog */
				  }
				  ;
				}
			      ;
			      /*^compute */
			      /*_.IFCPP___V61*/ meltfptr[58] =
				/*_._IFELSE___V62*/ meltfptr[61];;
			      /*epilog */

			      MELT_LOCATION
				("warmelt-normal.melt:6186:/ clear");
		    /*clear *//*_._IFELSE___V62*/ meltfptr[61] = 0;
			    }

#else /*MELT_HAVE_DEBUG */
			    /*^cppif.else */
			    /*_.IFCPP___V61*/ meltfptr[58] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
			    ;
			    MELT_LOCATION
			      ("warmelt-normal.melt:6185:/ quasiblock");


			    /*_.PROGN___V63*/ meltfptr[61] =
			      /*_.IFCPP___V61*/ meltfptr[58];;
			    /*^compute */
			    /*_._IFELSE___V53*/ meltfptr[51] =
			      /*_.PROGN___V63*/ meltfptr[61];;
			    /*epilog */

			    MELT_LOCATION
			      ("warmelt-normal.melt:6175:/ clear");
		  /*clear *//*_.IFCPP___V61*/ meltfptr[58] = 0;
			    /*^clear */
		  /*clear *//*_.PROGN___V63*/ meltfptr[61] = 0;
			  }
			  ;
			}
		      ;
		      /*_._IFELSE___V47*/ meltfptr[45] =
			/*_._IFELSE___V53*/ meltfptr[51];;
		      /*epilog */

		      MELT_LOCATION ("warmelt-normal.melt:6165:/ clear");
		/*clear *//*_#IS_LIST__L27*/ meltfnum[10] = 0;
		      /*^clear */
		/*clear *//*_._IFELSE___V53*/ meltfptr[51] = 0;
		    }
		    ;
		  }
		;
		/*_._IFELSE___V44*/ meltfptr[42] =
		  /*_._IFELSE___V47*/ meltfptr[45];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6160:/ clear");
	      /*clear *//*_#IS_MULTIPLE__L22*/ meltfnum[20] = 0;
		/*^clear */
	      /*clear *//*_._IFELSE___V47*/ meltfptr[45] = 0;
	      }
	      ;
	    }
	  ;
	  /*_.LET___V40*/ meltfptr[38] = /*_._IFELSE___V44*/ meltfptr[42];;

	  MELT_LOCATION ("warmelt-normal.melt:6157:/ clear");
	    /*clear *//*_.CURCONSBIND__V41*/ meltfptr[40] = 0;
	  /*^clear */
	    /*clear *//*_#IS_A__L19*/ meltfnum[9] = 0;
	  /*^clear */
	    /*clear *//*_._IFELSE___V44*/ meltfptr[42] = 0;
	  /*_.LET___V28*/ meltfptr[23] = /*_.LET___V40*/ meltfptr[38];;

	  MELT_LOCATION ("warmelt-normal.melt:6149:/ clear");
	    /*clear *//*_.LOCA_LOCATION__V29*/ meltfptr[28] = 0;
	  /*^clear */
	    /*clear *//*_.LOCB__V30*/ meltfptr[29] = 0;
	  /*^clear */
	    /*clear *//*_.CURSYMB__V32*/ meltfptr[30] = 0;
	  /*^clear */
	    /*clear *//*_.CUREXPR__V33*/ meltfptr[32] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V36*/ meltfptr[34] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V38*/ meltfptr[33] = 0;
	  /*^clear */
	    /*clear *//*_.LET___V40*/ meltfptr[38] = 0;
	  if ( /*_#SBIX__L9*/ meltfnum[8] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit1__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:6144:/ clear");
	    /*clear *//*_.CURSBIND__V23*/ meltfptr[22] = 0;
      /*^clear */
	    /*clear *//*_#SBIX__L9*/ meltfnum[8] = 0;
      /*^clear */
	    /*clear *//*_.IFCPP___V26*/ meltfptr[24] = 0;
      /*^clear */
	    /*clear *//*_.LET___V28*/ meltfptr[23] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6191:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L31*/ meltfnum[27] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6191:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L31*/ meltfnum[27])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L32*/ meltfnum[28] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6191:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L32*/ meltfnum[28];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6191;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec consbindlist";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CONSBINDLIST__V19*/ meltfptr[18];
	      /*_.MELT_DEBUG_FUN__V65*/ meltfptr[61] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V64*/ meltfptr[58] =
	      /*_.MELT_DEBUG_FUN__V65*/ meltfptr[61];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6191:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L32*/ meltfnum[28] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V65*/ meltfptr[61] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V64*/ meltfptr[58] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6191:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L31*/ meltfnum[27] = 0;
      /*^clear */
	     /*clear *//*_._IF___V64*/ meltfptr[58] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6192:/ quasiblock");


 /*_#NBCONSBIND__L33*/ meltfnum[10] =
      (melt_list_length
       ((melt_ptr_t) ( /*_.CONSBINDLIST__V19*/ meltfptr[18])));;
    /*^compute */
 /*_#CURCIX__L34*/ meltfnum[20] = 0;;
    /*^compute */
 /*_.NLOCSYMS__V67*/ meltfptr[45] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[13])),
	( /*_#NBCONSBIND__L33*/ meltfnum[10])));;
    /*^compute */
 /*_.NCBINDTUP__V68*/ meltfptr[40] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[13])),
	( /*_#NBCONSBIND__L33*/ meltfnum[10])));;
    /*^compute */
 /*_.RECFILLBINDINGS__V69*/ meltfptr[42] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[4]))));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6198:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NREP_LETREC */
					     meltfrout->tabval[14])), (6),
			      "CLASS_NREP_LETREC");
  /*_.INST__V71*/ meltfptr[29] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (0),
			  ( /*_.SLOC__V15*/ meltfptr[10]), "NREP_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLET_BINDINGS",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (1),
			  ( /*_.NCBINDTUP__V68*/ meltfptr[40]),
			  "NLET_BINDINGS");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLETREC_FILL_BINDINGS",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (3),
			  ( /*_.RECFILLBINDINGS__V69*/ meltfptr[42]),
			  "NLETREC_FILL_BINDINGS");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLETREC_LOCSYMS",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (5),
			  ( /*_.NLOCSYMS__V67*/ meltfptr[45]),
			  "NLETREC_LOCSYMS");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLET_BODY",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (2),
			  (( /*nil */ NULL)), "NLET_BODY");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NLETREC_BODY_BINDINGS",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V71*/ meltfptr[29])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V71*/ meltfptr[29]), (4),
			  (( /*nil */ NULL)), "NLETREC_BODY_BINDINGS");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V71*/ meltfptr[29],
				  "newly made instance");
    ;
    /*_.NLETREC__V70*/ meltfptr[28] = /*_.INST__V71*/ meltfptr[29];;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6208:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L35*/ meltfnum[9] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6208:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L35*/ meltfnum[9])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L36*/ meltfnum[28] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6208:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L36*/ meltfnum[28];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6208;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec unfilled nletrec";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V70*/ meltfptr[28];
	      /*_.MELT_DEBUG_FUN__V73*/ meltfptr[32] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V72*/ meltfptr[30] =
	      /*_.MELT_DEBUG_FUN__V73*/ meltfptr[32];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6208:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L36*/ meltfnum[28] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V73*/ meltfptr[32] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V72*/ meltfptr[30] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6208:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L35*/ meltfnum[9] = 0;
      /*^clear */
	     /*clear *//*_._IF___V72*/ meltfptr[30] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    /*citerblock FOREACH_IN_LIST */
    {
      /* start foreach_in_list meltcit4__EACHLIST */
      for ( /*_.CURPAIR__V74*/ meltfptr[34] =
	   melt_list_first ((melt_ptr_t) /*_.CONSBINDLIST__V19*/
			    meltfptr[18]);
	   melt_magic_discr ((melt_ptr_t) /*_.CURPAIR__V74*/ meltfptr[34]) ==
	   MELTOBMAG_PAIR;
	   /*_.CURPAIR__V74*/ meltfptr[34] =
	   melt_pair_tail ((melt_ptr_t) /*_.CURPAIR__V74*/ meltfptr[34]))
	{
	  /*_.CURCBIND__V75*/ meltfptr[33] =
	    melt_pair_head ((melt_ptr_t) /*_.CURPAIR__V74*/ meltfptr[34]);



#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6215:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L37*/ meltfnum[27] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6215:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L37*/ meltfnum[27])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L38*/ meltfnum[28] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6215:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L38*/ meltfnum[28];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6215;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normexp_letrec curcbind";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURCBIND__V75*/ meltfptr[33];
		    /*_.MELT_DEBUG_FUN__V77*/ meltfptr[61] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V76*/ meltfptr[38] =
		    /*_.MELT_DEBUG_FUN__V77*/ meltfptr[61];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6215:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L38*/ meltfnum[28] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V77*/ meltfptr[61] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V76*/ meltfptr[38] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6215:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L37*/ meltfnum[27] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V76*/ meltfptr[38] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6216:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L39*/ meltfnum[9] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURCBIND__V75*/ meltfptr[33]),
				   (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])));;
	    MELT_LOCATION ("warmelt-normal.melt:6216:/ cond");
	    /*cond */ if ( /*_#IS_A__L39*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V79*/ meltfptr[32] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6216:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curcbind"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6216) ? (6216) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V79*/ meltfptr[32] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V78*/ meltfptr[58] =
	      /*_._IFELSE___V79*/ meltfptr[32];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6216:/ clear");
	      /*clear *//*_#IS_A__L39*/ meltfnum[9] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V79*/ meltfptr[32] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V78*/ meltfptr[58] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6217:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.NCBINDTUP__V68*/ meltfptr[40]),
				     ( /*_#CURCIX__L34*/ meltfnum[20]),
				     (melt_ptr_t) ( /*_.CURCBIND__V75*/
						   meltfptr[33]));
	  }
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6218:/ locexp");
	    melt_put_int ((melt_ptr_t) ( /*_.CURCBIND__V75*/ meltfptr[33]),
			  ( /*_#CURCIX__L34*/ meltfnum[20]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6219:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURCBIND__V75*/
					       meltfptr[33]),
					      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		/*^putslot */
		/*putslot */
		melt_assertmsg ("checkobj putslot  _ @NCONSB_NLETREC",
				melt_magic_discr ((melt_ptr_t)
						  ( /*_.CURCBIND__V75*/
						   meltfptr[33])) ==
				MELTOBMAG_OBJECT);
		melt_putfield_object (( /*_.CURCBIND__V75*/ meltfptr[33]),
				      (3), ( /*_.NLETREC__V70*/ meltfptr[28]),
				      "NCONSB_NLETREC");
		;
		/*^touch */
		meltgc_touch ( /*_.CURCBIND__V75*/ meltfptr[33]);
		;
		/*^touchobj */

		melt_dbgtrace_written_object ( /*_.CURCBIND__V75*/
					      meltfptr[33], "put-fields");
		;
		/*epilog */
	      }
	      ;
	    }			/*noelse */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6220:/ quasiblock");


	  /*^getslot */
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURCBIND__V75*/ meltfptr[33]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 2, "NCONSB_DISCR");
   /*_.CURBDISCR__V80*/ meltfptr[30] = slot;
	  };
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6221:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURCBIND__V75*/ meltfptr[33]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 0, "BINDER");
   /*_.CURBINDER__V81*/ meltfptr[61] = slot;
	  };
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6223:/ cond");
	  /*cond */ if (
			 /*ifisa */
			 melt_is_instance_of ((melt_ptr_t)
					      ( /*_.CURCBIND__V75*/
					       meltfptr[33]),
					      (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])))
	    )			/*then */
	    {
	      /*^cond.then */
	      /*^getslot */
	      {
		melt_ptr_t slot = NULL, obj = NULL;
		obj =
		  (melt_ptr_t) ( /*_.CURCBIND__V75*/ meltfptr[33]) /*=obj*/ ;
		melt_object_get_field (slot, obj, 1, "NCONSB_LOC");
    /*_.NCONSB_LOC__V82*/ meltfptr[38] = slot;
	      };
	      ;
	    }
	  else
	    {			/*^cond.else */

   /*_.NCONSB_LOC__V82*/ meltfptr[38] = NULL;;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6223:/ cond");
	  /*cond */ if ( /*_.NCONSB_LOC__V82*/ meltfptr[38])	/*then */
	    {
	      /*^cond.then */
	      /*_.OR___V83*/ meltfptr[32] =
		/*_.NCONSB_LOC__V82*/ meltfptr[38];;
	    }
	  else
	    {
	      MELT_LOCATION ("warmelt-normal.melt:6223:/ cond.else");

	      /*^block */
	      /*anyblock */
	      {

		/*^cond */
		/*cond */ if ( /*_.SLOC__V15*/ meltfptr[10])	/*then */
		  {
		    /*^cond.then */
		    /*_.OR___V84*/ meltfptr[83] =
		      /*_.SLOC__V15*/ meltfptr[10];;
		  }
		else
		  {
		    MELT_LOCATION ("warmelt-normal.melt:6223:/ cond.else");

		    /*_.OR___V84*/ meltfptr[83] =
		      /*_.PSLOC__V5*/ meltfptr[4];;
		  }
		;
		/*^compute */
		/*_.OR___V83*/ meltfptr[32] = /*_.OR___V84*/ meltfptr[83];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6223:/ clear");
	      /*clear *//*_.OR___V84*/ meltfptr[83] = 0;
	      }
	      ;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6222:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_LOCSYMOCC */
						   meltfrout->tabval[15])),
				    (4), "CLASS_NREP_LOCSYMOCC");
   /*_.INST__V86*/ meltfptr[85] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V86*/ meltfptr[85]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V86*/ meltfptr[85]), (0),
				( /*_.OR___V83*/ meltfptr[32]), "NREP_LOC");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NOCC_CTYP",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V86*/ meltfptr[85]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V86*/ meltfptr[85]), (2),
				(( /*!CTYPE_VALUE */ meltfrout->tabval[16])),
				"NOCC_CTYP");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NOCC_SYMB",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V86*/ meltfptr[85]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V86*/ meltfptr[85]), (1),
				( /*_.CURBINDER__V81*/ meltfptr[61]),
				"NOCC_SYMB");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NOCC_BIND",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V86*/ meltfptr[85]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V86*/ meltfptr[85]), (3),
				( /*_.CURCBIND__V75*/ meltfptr[33]),
				"NOCC_BIND");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V86*/ meltfptr[85],
					"newly made instance");
	  ;
	  /*_.NLOCSYM__V85*/ meltfptr[83] = /*_.INST__V86*/ meltfptr[85];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6230:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.NLOCSYMS__V67*/ meltfptr[45]),
				     ( /*_#CURCIX__L34*/ meltfnum[20]),
				     (melt_ptr_t) ( /*_.NLOCSYM__V85*/
						   meltfptr[83]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
  /*_#NULL__L40*/ meltfnum[28] =
	    (( /*_.CURBDISCR__V80*/ meltfptr[30]) == NULL);;
	  MELT_LOCATION ("warmelt-normal.melt:6231:/ cond");
	  /*cond */ if ( /*_#NULL__L40*/ meltfnum[28])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		MELT_LOCATION ("warmelt-normal.melt:6232:/ quasiblock");


		/*^cond */
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.CURCBIND__V75*/
						     meltfptr[33]),
						    (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_INSTANCE_BINDING */ meltfrout->tabval[17])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.CURCBIND__V75*/ meltfptr[33])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 5, "NINSTB_CLABIND");
      /*_.CLABIND__V87*/ meltfptr[86] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

     /*_.CLABIND__V87*/ meltfptr[86] = NULL;;
		  }
		;
		MELT_LOCATION ("warmelt-normal.melt:6233:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.CLABIND__V87*/
						     meltfptr[86]),
						    (melt_ptr_t) (( /*!CLASS_ANY_BINDING */ meltfrout->tabval[18])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^getslot */
		    {
		      melt_ptr_t slot = NULL, obj = NULL;
		      obj =
			(melt_ptr_t) ( /*_.CLABIND__V87*/ meltfptr[86])
			/*=obj*/ ;
		      melt_object_get_field (slot, obj, 0, "BINDER");
      /*_.CLASYM__V88*/ meltfptr[87] = slot;
		    };
		    ;
		  }
		else
		  {		/*^cond.else */

     /*_.CLASYM__V88*/ meltfptr[87] = NULL;;
		  }
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6235:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#MELT_NEED_DBG__L41*/ meltfnum[27] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:6235:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L41*/ meltfnum[27])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	/*_#MELT_CALLCOUNT__L42*/ meltfnum[9] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:6235:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[7];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L42*/ meltfnum[9];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 6235;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring =
			    "normexp_letrec clabind=";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CLABIND__V87*/ meltfptr[86];
			  /*^apply.arg */
			  argtab[5].meltbp_cstring = " clasym=";
			  /*^apply.arg */
			  argtab[6].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CLASYM__V88*/ meltfptr[87];
			  /*_.MELT_DEBUG_FUN__V90*/ meltfptr[89] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[3])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V89*/ meltfptr[88] =
			  /*_.MELT_DEBUG_FUN__V90*/ meltfptr[89];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6235:/ clear");
		  /*clear *//*_#MELT_CALLCOUNT__L42*/ meltfnum[9] = 0;
			/*^clear */
		  /*clear *//*_.MELT_DEBUG_FUN__V90*/ meltfptr[89] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

       /*_._IF___V89*/ meltfptr[88] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:6235:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		/*clear *//*_#MELT_NEED_DBG__L41*/ meltfnum[27] = 0;
		  /*^clear */
		/*clear *//*_._IF___V89*/ meltfptr[88] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6237:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#IS_A__L43*/ meltfnum[9] =
		    melt_is_instance_of ((melt_ptr_t)
					 ( /*_.CURCBIND__V75*/ meltfptr[33]),
					 (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTED_INSTANCE_BINDING */ meltfrout->tabval[17])));;
		  MELT_LOCATION ("warmelt-normal.melt:6237:/ cond");
		  /*cond */ if ( /*_#IS_A__L43*/ meltfnum[9])	/*then */
		    {
		      /*^cond.then */
		      /*_._IFELSE___V92*/ meltfptr[88] = ( /*nil */ NULL);;
		    }
		  else
		    {
		      MELT_LOCATION ("warmelt-normal.melt:6237:/ cond.else");

		      /*^block */
		      /*anyblock */
		      {




			{
			  /*^locexp */
			  melt_assert_failed (("check curcbind for instance"),
					      ("warmelt-normal.melt")
					      ? ("warmelt-normal.melt") :
					      __FILE__,
					      (6237) ? (6237) : __LINE__,
					      __FUNCTION__);
			  ;
			}
			;
		  /*clear *//*_._IFELSE___V92*/ meltfptr[88] = 0;
			/*epilog */
		      }
		      ;
		    }
		  ;
		  /*^compute */
		  /*_.IFCPP___V91*/ meltfptr[89] =
		    /*_._IFELSE___V92*/ meltfptr[88];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6237:/ clear");
		/*clear *//*_#IS_A__L43*/ meltfnum[9] = 0;
		  /*^clear */
		/*clear *//*_._IFELSE___V92*/ meltfptr[88] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*_.IFCPP___V91*/ meltfptr[89] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
		;
		MELT_LOCATION ("warmelt-normal.melt:6238:/ quasiblock");



		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6239:/ msend");
		/*msend */
		{
		  union meltparam_un argtab[3];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^ojbmsend.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.ENV__V3*/ meltfptr[2];
		  /*^ojbmsend.arg */
		  argtab[1].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
		  /*^ojbmsend.arg */
		  argtab[2].meltbp_aptr =
		    (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[10];
		  /*_.CLADATA__V93*/ meltfptr[88] =
		    meltgc_send ((melt_ptr_t)
				 ( /*_.CLASYM__V88*/ meltfptr[87]),
				 (melt_ptr_t) (( /*!NORMAL_EXP */ meltfrout->
						tabval[19])),
				 (MELTBPARSTR_PTR MELTBPARSTR_PTR
				  MELTBPARSTR_PTR ""), argtab, "",
				 (union meltparam_un *) 0);
		}
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6241:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
		    melt_dbgcounter++;
#endif
		    ;
		  }
		  ;

		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#MELT_NEED_DBG__L44*/ meltfnum[27] =
		    /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
		    ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
		    0		/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
		    ;;
		  MELT_LOCATION ("warmelt-normal.melt:6241:/ cond");
		  /*cond */ if ( /*_#MELT_NEED_DBG__L44*/ meltfnum[27])	/*then */
		    {
		      /*^cond.then */
		      /*^block */
		      /*anyblock */
		      {

	/*_#MELT_CALLCOUNT__L45*/ meltfnum[9] =
			  /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
			  meltcallcount	/* melt_callcount debugging */
#else
			  0L	/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
			  ;;

			MELT_CHECK_SIGNAL ();
			;
			MELT_LOCATION ("warmelt-normal.melt:6241:/ apply");
			/*apply */
			{
			  union meltparam_un argtab[5];
			  memset (&argtab, 0, sizeof (argtab));
			  /*^apply.arg */
			  argtab[0].meltbp_long =
			    /*_#MELT_CALLCOUNT__L45*/ meltfnum[9];
			  /*^apply.arg */
			  argtab[1].meltbp_cstring = "warmelt-normal.melt";
			  /*^apply.arg */
			  argtab[2].meltbp_long = 6241;
			  /*^apply.arg */
			  argtab[3].meltbp_cstring = "normexp_letrec cladata";
			  /*^apply.arg */
			  argtab[4].meltbp_aptr =
			    (melt_ptr_t *) & /*_.CLADATA__V93*/ meltfptr[88];
			  /*_.MELT_DEBUG_FUN__V95*/ meltfptr[94] =
			    melt_apply ((meltclosure_ptr_t)
					(( /*!MELT_DEBUG_FUN */ meltfrout->
					  tabval[3])),
					(melt_ptr_t) (( /*nil */ NULL)),
					(MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_LONG MELTBPARSTR_CSTRING
					 MELTBPARSTR_PTR ""), argtab, "",
					(union meltparam_un *) 0);
			}
			;
			/*_._IF___V94*/ meltfptr[93] =
			  /*_.MELT_DEBUG_FUN__V95*/ meltfptr[94];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6241:/ clear");
		  /*clear *//*_#MELT_CALLCOUNT__L45*/ meltfnum[9] = 0;
			/*^clear */
		  /*clear *//*_.MELT_DEBUG_FUN__V95*/ meltfptr[94] = 0;
		      }
		      ;
		    }
		  else
		    {		/*^cond.else */

       /*_._IF___V94*/ meltfptr[93] = NULL;;
		    }
		  ;

		  {
		    MELT_LOCATION ("warmelt-normal.melt:6241:/ locexp");
		    /*void */ (void) 0;
		  }
		  ;
		  /*^quasiblock */


		  /*epilog */

		  /*^clear */
		/*clear *//*_#MELT_NEED_DBG__L44*/ meltfnum[27] = 0;
		  /*^clear */
		/*clear *//*_._IF___V94*/ meltfptr[93] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*^block */
		/*anyblock */
		{


		  {
		    /*^locexp */
		    /*void */ (void) 0;
		  }
		  ;
		  /*epilog */
		}

#endif /*MELT_HAVE_DEBUG */
		;

#if MELT_HAVE_DEBUG
		MELT_LOCATION ("warmelt-normal.melt:6242:/ cppif.then");
		/*^block */
		/*anyblock */
		{


		  MELT_CHECK_SIGNAL ();
		  ;
      /*_#IS_A__L46*/ meltfnum[9] =
		    melt_is_instance_of ((melt_ptr_t)
					 ( /*_.CLADATA__V93*/ meltfptr[88]),
					 (melt_ptr_t) (( /*!CLASS_NREP */
							meltfrout->
							tabval[20])));;
		  MELT_LOCATION ("warmelt-normal.melt:6242:/ cond");
		  /*cond */ if ( /*_#IS_A__L46*/ meltfnum[9])	/*then */
		    {
		      /*^cond.then */
		      /*_._IFELSE___V97*/ meltfptr[93] = ( /*nil */ NULL);;
		    }
		  else
		    {
		      MELT_LOCATION ("warmelt-normal.melt:6242:/ cond.else");

		      /*^block */
		      /*anyblock */
		      {




			{
			  /*^locexp */
			  melt_assert_failed (("check cladata"),
					      ("warmelt-normal.melt")
					      ? ("warmelt-normal.melt") :
					      __FILE__,
					      (6242) ? (6242) : __LINE__,
					      __FUNCTION__);
			  ;
			}
			;
		  /*clear *//*_._IFELSE___V97*/ meltfptr[93] = 0;
			/*epilog */
		      }
		      ;
		    }
		  ;
		  /*^compute */
		  /*_.IFCPP___V96*/ meltfptr[94] =
		    /*_._IFELSE___V97*/ meltfptr[93];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6242:/ clear");
		/*clear *//*_#IS_A__L46*/ meltfnum[9] = 0;
		  /*^clear */
		/*clear *//*_._IFELSE___V97*/ meltfptr[93] = 0;
		}

#else /*MELT_HAVE_DEBUG */
		/*^cppif.else */
		/*_.IFCPP___V96*/ meltfptr[94] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6243:/ cond");
		/*cond */ if (
			       /*ifisa */
			       melt_is_instance_of ((melt_ptr_t)
						    ( /*_.CURCBIND__V75*/
						     meltfptr[33]),
						    (melt_ptr_t) (( /*!CLASS_NORMAL_CONSTRUCTOR_BINDING */ meltfrout->tabval[11])))
		  )		/*then */
		  {
		    /*^cond.then */
		    /*^block */
		    /*anyblock */
		    {

		      /*^putslot */
		      /*putslot */
		      melt_assertmsg ("checkobj putslot  _ @NCONSB_DISCR",
				      melt_magic_discr ((melt_ptr_t)
							( /*_.CURCBIND__V75*/
							 meltfptr[33])) ==
				      MELTOBMAG_OBJECT);
		      melt_putfield_object (( /*_.CURCBIND__V75*/
					     meltfptr[33]), (2),
					    ( /*_.CLADATA__V93*/
					     meltfptr[88]), "NCONSB_DISCR");
		      ;
		      /*^touch */
		      meltgc_touch ( /*_.CURCBIND__V75*/ meltfptr[33]);
		      ;
		      /*^touchobj */

		      melt_dbgtrace_written_object ( /*_.CURCBIND__V75*/
						    meltfptr[33],
						    "put-fields");
		      ;
		      /*epilog */
		    }
		    ;
		  }		/*noelse */
		;

		MELT_LOCATION ("warmelt-normal.melt:6238:/ clear");
	      /*clear *//*_.CLADATA__V93*/ meltfptr[88] = 0;
		/*^clear */
	      /*clear *//*_.IFCPP___V96*/ meltfptr[94] = 0;

		MELT_LOCATION ("warmelt-normal.melt:6232:/ clear");
	      /*clear *//*_.CLABIND__V87*/ meltfptr[86] = 0;
		/*^clear */
	      /*clear *//*_.CLASYM__V88*/ meltfptr[87] = 0;
		/*^clear */
	      /*clear *//*_.IFCPP___V91*/ meltfptr[89] = 0;
		/*epilog */
	      }
	      ;
	    }			/*noelse */
	  ;

	  MELT_LOCATION ("warmelt-normal.melt:6220:/ clear");
	    /*clear *//*_.CURBDISCR__V80*/ meltfptr[30] = 0;
	  /*^clear */
	    /*clear *//*_.CURBINDER__V81*/ meltfptr[61] = 0;
	  /*^clear */
	    /*clear *//*_.NCONSB_LOC__V82*/ meltfptr[38] = 0;
	  /*^clear */
	    /*clear *//*_.OR___V83*/ meltfptr[32] = 0;
	  /*^clear */
	    /*clear *//*_.NLOCSYM__V85*/ meltfptr[83] = 0;
	  /*^clear */
	    /*clear *//*_#NULL__L40*/ meltfnum[28] = 0;
  /*_#plI__L47*/ meltfnum[27] =
	    (( /*_#CURCIX__L34*/ meltfnum[20]) + (1));;
	  MELT_LOCATION ("warmelt-normal.melt:6246:/ compute");
	  /*_#CURCIX__L34*/ meltfnum[20] = /*_#SETQ___L48*/ meltfnum[9] =
	    /*_#plI__L47*/ meltfnum[27];;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6247:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.CURCBIND__V75*/ meltfptr[33];
	    /*_.PUT_ENV__V98*/ meltfptr[93] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!PUT_ENV */ meltfrout->tabval[21])),
			  (melt_ptr_t) ( /*_.NEWENV__V20*/ meltfptr[19]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;
	}			/* end foreach_in_list meltcit4__EACHLIST */
     /*_.CURPAIR__V74*/ meltfptr[34] = NULL;
     /*_.CURCBIND__V75*/ meltfptr[33] = NULL;


      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:6212:/ clear");
	    /*clear *//*_.CURPAIR__V74*/ meltfptr[34] = 0;
      /*^clear */
	    /*clear *//*_.CURCBIND__V75*/ meltfptr[33] = 0;
      /*^clear */
	    /*clear *//*_.IFCPP___V78*/ meltfptr[58] = 0;
      /*^clear */
	    /*clear *//*_#plI__L47*/ meltfnum[27] = 0;
      /*^clear */
	    /*clear *//*_#SETQ___L48*/ meltfnum[9] = 0;
      /*^clear */
	    /*clear *//*_.PUT_ENV__V98*/ meltfptr[93] = 0;
    }				/*endciterblock FOREACH_IN_LIST */
    ;
    /*citerblock FOREACH_IN_MULTIPLE */
    {
      /* start foreach_in_multiple meltcit5__EACHTUP */
      long meltcit5__EACHTUP_ln =
	melt_multiple_length ((melt_ptr_t) /*_.SBINDS__V16*/ meltfptr[15]);
      for ( /*_#SBIX__L49*/ meltfnum[28] = 0;
	   ( /*_#SBIX__L49*/ meltfnum[28] >= 0)
	   && ( /*_#SBIX__L49*/ meltfnum[28] < meltcit5__EACHTUP_ln);
	/*_#SBIX__L49*/ meltfnum[28]++)
	{
	  /*_.CURSBIND__V99*/ meltfptr[88] =
	    melt_multiple_nth ((melt_ptr_t) ( /*_.SBINDS__V16*/ meltfptr[15]),
			       /*_#SBIX__L49*/ meltfnum[28]);




#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6254:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L50*/ meltfnum[49] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6254:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L50*/ meltfnum[49])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L51*/ meltfnum[50] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6254:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[7];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L51*/ meltfnum[50];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6254;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normexp_letrec thirdloop cursbind=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURSBIND__V99*/ meltfptr[88];
		    /*^apply.arg */
		    argtab[5].meltbp_cstring = " sbix=";
		    /*^apply.arg */
		    argtab[6].meltbp_long = /*_#SBIX__L49*/ meltfnum[28];
		    /*_.MELT_DEBUG_FUN__V101*/ meltfptr[86] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V100*/ meltfptr[94] =
		    /*_.MELT_DEBUG_FUN__V101*/ meltfptr[86];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6254:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L51*/ meltfnum[50] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V101*/ meltfptr[86] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V100*/ meltfptr[94] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6254:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L50*/ meltfnum[49] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V100*/ meltfptr[94] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6255:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L52*/ meltfnum[50] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.CURSBIND__V99*/ meltfptr[88]),
				   (melt_ptr_t) (( /*!CLASS_SOURCE_LETREC_BINDING */ meltfrout->tabval[7])));;
	    MELT_LOCATION ("warmelt-normal.melt:6255:/ cond");
	    /*cond */ if ( /*_#IS_A__L52*/ meltfnum[50])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V103*/ meltfptr[89] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6255:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check cursbind"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6255) ? (6255) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V103*/ meltfptr[89] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V102*/ meltfptr[87] =
	      /*_._IFELSE___V103*/ meltfptr[89];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6255:/ clear");
	      /*clear *//*_#IS_A__L52*/ meltfnum[50] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V103*/ meltfptr[89] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V102*/ meltfptr[87] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6256:/ quasiblock");


	  /*^getslot */
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V99*/ meltfptr[88]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
   /*_.LOCA_LOCATION__V104*/ meltfptr[30] = slot;
	  };
	  ;
	  /*^cond */
	  /*cond */ if ( /*_.LOCA_LOCATION__V104*/ meltfptr[30])	/*then */
	    {
	      /*^cond.then */
	      /*_.LOCB__V105*/ meltfptr[61] =
		/*_.LOCA_LOCATION__V104*/ meltfptr[30];;
	    }
	  else
	    {
	      MELT_LOCATION ("warmelt-normal.melt:6256:/ cond.else");

	      /*^block */
	      /*anyblock */
	      {

		/*^cond */
		/*cond */ if ( /*_.SLOC__V15*/ meltfptr[10])	/*then */
		  {
		    /*^cond.then */
		    /*_.OR___V106*/ meltfptr[38] =
		      /*_.SLOC__V15*/ meltfptr[10];;
		  }
		else
		  {
		    MELT_LOCATION ("warmelt-normal.melt:6256:/ cond.else");

		    /*_.OR___V106*/ meltfptr[38] =
		      /*_.PSLOC__V5*/ meltfptr[4];;
		  }
		;
		/*^compute */
		/*_.LOCB__V105*/ meltfptr[61] = /*_.OR___V106*/ meltfptr[38];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6256:/ clear");
	      /*clear *//*_.OR___V106*/ meltfptr[38] = 0;
	      }
	      ;
	    }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6257:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V99*/ meltfptr[88]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 3, "SLETB_BINDER");
   /*_.CURSYMB__V107*/ meltfptr[32] = slot;
	  };
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6258:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURSBIND__V99*/ meltfptr[88]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 4, "SLETB_EXPR");
   /*_.CUREXPR__V108*/ meltfptr[83] = slot;
	  };
	  ;
  /*_.CURBIND__V109*/ meltfptr[86] =
	    /*mapobject_get */
	    melt_get_mapobjects ((meltmapobjects_ptr_t)
				 ( /*_.SYMBINDMAP__V22*/ meltfptr[21]),
				 (meltobject_ptr_t) ( /*_.CURSYMB__V107*/
						     meltfptr[32]));;
	  /*^compute */
  /*_.CURSEXPR__V110*/ meltfptr[94] =
	    /*mapobject_get */
	    melt_get_mapobjects ((meltmapobjects_ptr_t)
				 ( /*_.SYMBEXPRMAP__V21*/ meltfptr[20]),
				 (meltobject_ptr_t) ( /*_.CURSYMB__V107*/
						     meltfptr[32]));;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6262:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L53*/ meltfnum[49] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6262:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L53*/ meltfnum[49])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L54*/ meltfnum[50] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6262:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[11];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L54*/ meltfnum[50];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6262;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normexp_letrec thirdloop curexpr=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CUREXPR__V108*/ meltfptr[83];
		    /*^apply.arg */
		    argtab[5].meltbp_cstring = " cursymb=";
		    /*^apply.arg */
		    argtab[6].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURSYMB__V107*/ meltfptr[32];
		    /*^apply.arg */
		    argtab[7].meltbp_cstring = " curbind=";
		    /*^apply.arg */
		    argtab[8].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURBIND__V109*/ meltfptr[86];
		    /*^apply.arg */
		    argtab[9].meltbp_cstring = " cursexpr=";
		    /*^apply.arg */
		    argtab[10].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CURSEXPR__V110*/ meltfptr[94];
		    /*_.MELT_DEBUG_FUN__V112*/ meltfptr[38] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V111*/ meltfptr[89] =
		    /*_.MELT_DEBUG_FUN__V112*/ meltfptr[38];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6262:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L54*/ meltfnum[50] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V112*/ meltfptr[38] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V111*/ meltfptr[89] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6262:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L53*/ meltfnum[49] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V111*/ meltfptr[89] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6266:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#eqeq__L55*/ meltfnum[50] =
	      (( /*_.CUREXPR__V108*/ meltfptr[83]) ==
	       ( /*_.CURSEXPR__V110*/ meltfptr[94]));;
	    MELT_LOCATION ("warmelt-normal.melt:6266:/ cond");
	    /*cond */ if ( /*_#eqeq__L55*/ meltfnum[50])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V114*/ meltfptr[89] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6266:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("check curxpr same cursexpr"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6266) ? (6266) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V114*/ meltfptr[89] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V113*/ meltfptr[38] =
	      /*_._IFELSE___V114*/ meltfptr[89];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6266:/ clear");
	      /*clear *//*_#eqeq__L55*/ meltfnum[50] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V114*/ meltfptr[89] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V113*/ meltfptr[38] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6267:/ msend");
	  /*msend */
	  {
	    union meltparam_un argtab[5];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^ojbmsend.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.CURSYMB__V107*/ meltfptr[32];
	    /*^ojbmsend.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.CURBIND__V109*/ meltfptr[86];
	    /*^ojbmsend.arg */
	    argtab[2].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NEWENV__V20*/ meltfptr[19];
	    /*^ojbmsend.arg */
	    argtab[3].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
	    /*^ojbmsend.arg */
	    argtab[4].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[10];
	    /*_.NORMAL_LETREC_CONSTRUCTIVE__V115*/ meltfptr[89] =
	      meltgc_send ((melt_ptr_t) ( /*_.CUREXPR__V108*/ meltfptr[83]),
			   (melt_ptr_t) (( /*!NORMAL_LETREC_CONSTRUCTIVE */
					  meltfrout->tabval[22])),
			   (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR
			    MELTBPARSTR_PTR MELTBPARSTR_PTR ""), argtab, "",
			   (union meltparam_un *) 0);
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6268:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L56*/ meltfnum[49] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6268:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L56*/ meltfnum[49])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L57*/ meltfnum[50] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6268:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L57*/ meltfnum[50];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6268;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normexp_letrec thirdloop done curexpr";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.CUREXPR__V108*/ meltfptr[83];
		    /*_.MELT_DEBUG_FUN__V117*/ meltfptr[116] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V116*/ meltfptr[115] =
		    /*_.MELT_DEBUG_FUN__V117*/ meltfptr[116];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6268:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L57*/ meltfnum[50] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V117*/ meltfptr[116] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V116*/ meltfptr[115] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6268:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L56*/ meltfnum[49] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V116*/ meltfptr[115] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_LOCATION ("warmelt-normal.melt:6256:/ clear");
	    /*clear *//*_.LOCA_LOCATION__V104*/ meltfptr[30] = 0;
	  /*^clear */
	    /*clear *//*_.LOCB__V105*/ meltfptr[61] = 0;
	  /*^clear */
	    /*clear *//*_.CURSYMB__V107*/ meltfptr[32] = 0;
	  /*^clear */
	    /*clear *//*_.CUREXPR__V108*/ meltfptr[83] = 0;
	  /*^clear */
	    /*clear *//*_.CURBIND__V109*/ meltfptr[86] = 0;
	  /*^clear */
	    /*clear *//*_.CURSEXPR__V110*/ meltfptr[94] = 0;
	  /*^clear */
	    /*clear *//*_.IFCPP___V113*/ meltfptr[38] = 0;
	  /*^clear */
	    /*clear *//*_.NORMAL_LETREC_CONSTRUCTIVE__V115*/ meltfptr[89] = 0;
	  if ( /*_#SBIX__L49*/ meltfnum[28] < 0)
	    break;
	}			/* end  foreach_in_multiple meltcit5__EACHTUP */

      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:6251:/ clear");
	    /*clear *//*_.CURSBIND__V99*/ meltfptr[88] = 0;
      /*^clear */
	    /*clear *//*_#SBIX__L49*/ meltfnum[28] = 0;
      /*^clear */
	    /*clear *//*_.IFCPP___V102*/ meltfptr[87] = 0;
    }				/*endciterblock FOREACH_IN_MULTIPLE */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6271:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L58*/ meltfnum[50] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6271:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L58*/ meltfnum[50])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L59*/ meltfnum[49] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6271:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L59*/ meltfnum[49];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6271;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring =
		"normexp_letrec recfillbindings before normalizing the body";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECFILLBINDINGS__V69*/ meltfptr[42];
	      /*_.MELT_DEBUG_FUN__V119*/ meltfptr[115] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V118*/ meltfptr[116] =
	      /*_.MELT_DEBUG_FUN__V119*/ meltfptr[115];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6271:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L59*/ meltfnum[49] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V119*/ meltfptr[115] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V118*/ meltfptr[116] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6271:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L58*/ meltfnum[50] = 0;
      /*^clear */
	     /*clear *//*_._IF___V118*/ meltfptr[116] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6274:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L60*/ meltfnum[49] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6274:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L60*/ meltfnum[49])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L61*/ meltfnum[50] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6274:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L61*/ meltfnum[50];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6274;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec normalizing sbody";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SBODY__V17*/ meltfptr[16];
	      /*_.MELT_DEBUG_FUN__V121*/ meltfptr[61] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V120*/ meltfptr[30] =
	      /*_.MELT_DEBUG_FUN__V121*/ meltfptr[61];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6274:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L61*/ meltfnum[50] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V121*/ meltfptr[61] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V120*/ meltfptr[30] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6274:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L60*/ meltfnum[49] = 0;
      /*^clear */
	     /*clear *//*_._IF___V120*/ meltfptr[30] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6275:/ quasiblock");


    /*^multiapply */
    /*multiapply 4args, 1x.res */
    {
      union meltparam_un argtab[3];

      union meltparam_un restab[1];
      memset (&restab, 0, sizeof (restab));
      memset (&argtab, 0, sizeof (argtab));
      /*^multiapply.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NEWENV__V20*/ meltfptr[19];	/*^multiapply.arg */
      argtab[1].meltbp_aptr = (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];	/*^multiapply.arg */
      argtab[2].meltbp_aptr = (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[10];
      /*^multiapply.xres */
      restab[0].meltbp_aptr =
	(melt_ptr_t *) & /*_.NBODBINDINGS__V124*/ meltfptr[86];
      /*^multiapply.appl */
      /*_.NBODY__V123*/ meltfptr[83] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!NORMALIZE_TUPLE */ meltfrout->tabval[23])),
		    (melt_ptr_t) ( /*_.SBODY__V17*/ meltfptr[16]),
		    (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_PTR ""),
		    argtab, (MELTBPARSTR_PTR ""), restab);
    }
    ;
    /*^quasiblock */



#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6278:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L62*/ meltfnum[50] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6278:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L62*/ meltfnum[50])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L63*/ meltfnum[49] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6278:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L63*/ meltfnum[49];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6278;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec nbody=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBODY__V123*/ meltfptr[83];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nbodbindings=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NBODBINDINGS__V124*/ meltfptr[86];
	      /*_.MELT_DEBUG_FUN__V126*/ meltfptr[38] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V125*/ meltfptr[94] =
	      /*_.MELT_DEBUG_FUN__V126*/ meltfptr[38];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6278:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L63*/ meltfnum[49] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V126*/ meltfptr[38] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V125*/ meltfptr[94] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6278:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L62*/ meltfnum[50] = 0;
      /*^clear */
	     /*clear *//*_._IF___V125*/ meltfptr[94] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6279:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L64*/ meltfnum[49] =
	(( /*_.NBODBINDINGS__V124*/ meltfptr[86]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.NBODBINDINGS__V124*/ meltfptr[86])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:6279:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L64*/ meltfnum[49])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V128*/ meltfptr[115] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6279:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("normexp_letrec check nbodbindings"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6279) ? (6279) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V128*/ meltfptr[115] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V127*/ meltfptr[89] = /*_._IFELSE___V128*/ meltfptr[115];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6279:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L64*/ meltfnum[49] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V128*/ meltfptr[115] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V127*/ meltfptr[89] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6280:/ cond");
    /*cond */ if (
		   /*ifisa */
		   melt_is_instance_of ((melt_ptr_t)
					( /*_.NLETREC__V70*/ meltfptr[28]),
					(melt_ptr_t) (( /*!CLASS_NREP_LETREC */ meltfrout->tabval[14])))
      )				/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NLET_BODY",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.NLETREC__V70*/
					     meltfptr[28])) ==
			  MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.NLETREC__V70*/ meltfptr[28]), (2),
				( /*_.NBODY__V123*/ meltfptr[83]),
				"NLET_BODY");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NLETREC_BODY_BINDINGS",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.NLETREC__V70*/
					     meltfptr[28])) ==
			  MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.NLETREC__V70*/ meltfptr[28]), (4),
				( /*_.NBODBINDINGS__V124*/ meltfptr[86]),
				"NLETREC_BODY_BINDINGS");
	  ;
	  /*^touch */
	  meltgc_touch ( /*_.NLETREC__V70*/ meltfptr[28]);
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.NLETREC__V70*/ meltfptr[28],
					"put-fields");
	  ;
	  /*epilog */
	}
	;
      }				/*noelse */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6283:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L65*/ meltfnum[50] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6283:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L65*/ meltfnum[50])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L66*/ meltfnum[49] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6283:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L66*/ meltfnum[49];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6283;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec updated nletrec";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V70*/ meltfptr[28];
	      /*_.MELT_DEBUG_FUN__V130*/ meltfptr[61] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V129*/ meltfptr[116] =
	      /*_.MELT_DEBUG_FUN__V130*/ meltfptr[61];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6283:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L66*/ meltfnum[49] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V130*/ meltfptr[61] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V129*/ meltfptr[116] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6283:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L65*/ meltfnum[50] = 0;
      /*^clear */
	     /*clear *//*_._IF___V129*/ meltfptr[116] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6284:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_LIST_OR_NULL__L67*/ meltfnum[49] =
	(( /*_.BINDLIST__V18*/ meltfptr[17]) == NULL
	 ||
	 (melt_unsafe_magic_discr
	  ((melt_ptr_t) ( /*_.BINDLIST__V18*/ meltfptr[17])) ==
	  MELTOBMAG_LIST));;
      MELT_LOCATION ("warmelt-normal.melt:6284:/ cond");
      /*cond */ if ( /*_#IS_LIST_OR_NULL__L67*/ meltfnum[49])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V132*/ meltfptr[38] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6284:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("normexp_letrec check bindlist"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6284) ? (6284) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V132*/ meltfptr[38] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V131*/ meltfptr[30] = /*_._IFELSE___V132*/ meltfptr[38];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6284:/ clear");
	     /*clear *//*_#IS_LIST_OR_NULL__L67*/ meltfnum[49] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V132*/ meltfptr[38] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V131*/ meltfptr[30] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6286:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 6, "NCTX_SYMBCACHEMAP");
  /*_.SYCMAP__V133*/ meltfptr[94] = slot;
    };
    ;
    /*citerblock FOREACH_IN_LIST */
    {
      /* start foreach_in_list meltcit6__EACHLIST */
      for ( /*_.CURPAIR__V134*/ meltfptr[115] =
	   melt_list_first ((melt_ptr_t) /*_.BINDLIST__V18*/ meltfptr[17]);
	   melt_magic_discr ((melt_ptr_t) /*_.CURPAIR__V134*/ meltfptr[115])
	   == MELTOBMAG_PAIR;
	   /*_.CURPAIR__V134*/ meltfptr[115] =
	   melt_pair_tail ((melt_ptr_t) /*_.CURPAIR__V134*/ meltfptr[115]))
	{
	  /*_.BND__V135*/ meltfptr[61] =
	    melt_pair_head ((melt_ptr_t) /*_.CURPAIR__V134*/ meltfptr[115]);



#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6290:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#MELT_NEED_DBG__L68*/ meltfnum[50] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6290:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L68*/ meltfnum[50])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

      /*_#MELT_CALLCOUNT__L69*/ meltfnum[49] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6290:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L69*/ meltfnum[49];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6290;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normexp_letrec removing from sycmap bnd";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.BND__V135*/ meltfptr[61];
		    /*_.MELT_DEBUG_FUN__V137*/ meltfptr[38] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[3])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V136*/ meltfptr[116] =
		    /*_.MELT_DEBUG_FUN__V137*/ meltfptr[38];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6290:/ clear");
		/*clear *//*_#MELT_CALLCOUNT__L69*/ meltfnum[49] = 0;
		  /*^clear */
		/*clear *//*_.MELT_DEBUG_FUN__V137*/ meltfptr[38] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

     /*_._IF___V136*/ meltfptr[116] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6290:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	      /*clear *//*_#MELT_NEED_DBG__L68*/ meltfnum[50] = 0;
	    /*^clear */
	      /*clear *//*_._IF___V136*/ meltfptr[116] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6291:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    MELT_CHECK_SIGNAL ();
	    ;
    /*_#IS_A__L70*/ meltfnum[49] =
	      melt_is_instance_of ((melt_ptr_t)
				   ( /*_.BND__V135*/ meltfptr[61]),
				   (melt_ptr_t) (( /*!CLASS_NORMAL_LET_BINDING */ meltfrout->tabval[24])));;
	    MELT_LOCATION ("warmelt-normal.melt:6291:/ cond");
	    /*cond */ if ( /*_#IS_A__L70*/ meltfnum[49])	/*then */
	      {
		/*^cond.then */
		/*_._IFELSE___V139*/ meltfptr[116] = ( /*nil */ NULL);;
	      }
	    else
	      {
		MELT_LOCATION ("warmelt-normal.melt:6291:/ cond.else");

		/*^block */
		/*anyblock */
		{




		  {
		    /*^locexp */
		    melt_assert_failed (("normexp_letrec check bnd"),
					("warmelt-normal.melt")
					? ("warmelt-normal.melt") : __FILE__,
					(6291) ? (6291) : __LINE__,
					__FUNCTION__);
		    ;
		  }
		  ;
		/*clear *//*_._IFELSE___V139*/ meltfptr[116] = 0;
		  /*epilog */
		}
		;
	      }
	    ;
	    /*^compute */
	    /*_.IFCPP___V138*/ meltfptr[38] =
	      /*_._IFELSE___V139*/ meltfptr[116];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6291:/ clear");
	      /*clear *//*_#IS_A__L70*/ meltfnum[49] = 0;
	    /*^clear */
	      /*clear *//*_._IFELSE___V139*/ meltfptr[116] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*_.IFCPP___V138*/ meltfptr[38] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6292:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.BND__V135*/ meltfptr[61]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 0, "BINDER");
   /*_.BINDER__V140*/ meltfptr[116] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    meltgc_remove_mapobjects ((meltmapobjects_ptr_t)
				      ( /*_.SYCMAP__V133*/ meltfptr[94]),
				      (meltobject_ptr_t) ( /*_.BINDER__V140*/
							  meltfptr[116]));
	  }
	  ;
	}			/* end foreach_in_list meltcit6__EACHLIST */
     /*_.CURPAIR__V134*/ meltfptr[115] = NULL;
     /*_.BND__V135*/ meltfptr[61] = NULL;


      /*citerepilog */

      MELT_LOCATION ("warmelt-normal.melt:6287:/ clear");
	    /*clear *//*_.CURPAIR__V134*/ meltfptr[115] = 0;
      /*^clear */
	    /*clear *//*_.BND__V135*/ meltfptr[61] = 0;
      /*^clear */
	    /*clear *//*_.IFCPP___V138*/ meltfptr[38] = 0;
      /*^clear */
	    /*clear *//*_.BINDER__V140*/ meltfptr[116] = 0;
    }				/*endciterblock FOREACH_IN_LIST */
    ;

    MELT_LOCATION ("warmelt-normal.melt:6286:/ clear");
	   /*clear *//*_.SYCMAP__V133*/ meltfptr[94] = 0;
    MELT_LOCATION ("warmelt-normal.melt:6295:/ quasiblock");


 /*_.NLASTBODY__V142*/ meltfptr[141] =
      (melt_multiple_nth
       ((melt_ptr_t) ( /*_.NBODY__V123*/ meltfptr[83]), (-1)));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6298:/ msend");
    /*msend */
    {
      union meltparam_un argtab[1];
      memset (&argtab, 0, sizeof (argtab));
      /*^ojbmsend.arg */
      argtab[0].meltbp_aptr = (melt_ptr_t *) & /*_.NEWENV__V20*/ meltfptr[19];
      /*_.GET_CTYPE__V143*/ meltfptr[142] =
	meltgc_send ((melt_ptr_t) ( /*_.NLASTBODY__V142*/ meltfptr[141]),
		     (melt_ptr_t) (( /*!GET_CTYPE */ meltfrout->tabval[26])),
		     (MELTBPARSTR_PTR ""), argtab, "",
		     (union meltparam_un *) 0);
    }
    ;
    /*^cond */
    /*cond */ if ( /*_.GET_CTYPE__V143*/ meltfptr[142])	/*then */
      {
	/*^cond.then */
	/*_.NLASTYP__V144*/ meltfptr[143] =
	  /*_.GET_CTYPE__V143*/ meltfptr[142];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6298:/ cond.else");

	/*_.NLASTYP__V144*/ meltfptr[143] =
	  ( /*!CTYPE_VOID */ meltfrout->tabval[25]);;
      }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6299:/ apply");
    /*apply */
    {
      /*_.CSYM__V145*/ meltfptr[144] =
	melt_apply ((meltclosure_ptr_t)
		    (( /*!CLONE_SYMBOL */ meltfrout->tabval[27])),
		    (melt_ptr_t) (( /*!konst_28_LETREC_ */ meltfrout->
				   tabval[28])), (""),
		    (union meltparam_un *) 0, "", (union meltparam_un *) 0);
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6300:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NORMAL_LET_BINDING */
					     meltfrout->tabval[24])), (4),
			      "CLASS_NORMAL_LET_BINDING");
  /*_.INST__V147*/ meltfptr[146] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @BINDER",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V147*/ meltfptr[146])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V147*/ meltfptr[146]), (0),
			  ( /*_.CSYM__V145*/ meltfptr[144]), "BINDER");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @LETBIND_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V147*/ meltfptr[146])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V147*/ meltfptr[146]), (3),
			  ( /*_.SLOC__V15*/ meltfptr[10]), "LETBIND_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @LETBIND_TYPE",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V147*/ meltfptr[146])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V147*/ meltfptr[146]), (1),
			  ( /*_.NLASTYP__V144*/ meltfptr[143]),
			  "LETBIND_TYPE");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @LETBIND_EXPR",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V147*/ meltfptr[146])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V147*/ meltfptr[146]), (2),
			  ( /*_.NLETREC__V70*/ meltfptr[28]), "LETBIND_EXPR");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V147*/ meltfptr[146],
				  "newly made instance");
    ;
    /*_.CBIND__V146*/ meltfptr[145] = /*_.INST__V147*/ meltfptr[146];;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6305:/ quasiblock");


    /*^rawallocobj */
    /*rawallocobj */
    {
      melt_ptr_t newobj = 0;
      melt_raw_object_create (newobj,
			      (melt_ptr_t) (( /*!CLASS_NREP_LOCSYMOCC */
					     meltfrout->tabval[15])), (4),
			      "CLASS_NREP_LOCSYMOCC");
  /*_.INST__V149*/ meltfptr[148] =
	newobj;
    };
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V149*/ meltfptr[148])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V149*/ meltfptr[148]), (0),
			  ( /*_.SLOC__V15*/ meltfptr[10]), "NREP_LOC");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NOCC_CTYP",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V149*/ meltfptr[148])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V149*/ meltfptr[148]), (2),
			  ( /*_.NLASTYP__V144*/ meltfptr[143]), "NOCC_CTYP");
    ;
    /*^putslot */
    /*putslot */
    melt_assertmsg ("checkobj putslot  _ @NOCC_BIND",
		    melt_magic_discr ((melt_ptr_t)
				      ( /*_.INST__V149*/ meltfptr[148])) ==
		    MELTOBMAG_OBJECT);
    melt_putfield_object (( /*_.INST__V149*/ meltfptr[148]), (3),
			  ( /*_.CBIND__V146*/ meltfptr[145]), "NOCC_BIND");
    ;
    /*^touchobj */

    melt_dbgtrace_written_object ( /*_.INST__V149*/ meltfptr[148],
				  "newly made instance");
    ;
    /*_.CLOCC__V148*/ meltfptr[147] = /*_.INST__V149*/ meltfptr[148];;
    /*^compute */
 /*_.RESBINDS__V150*/ meltfptr[149] =
      (meltgc_new_list
       ((meltobject_ptr_t) (( /*!DISCR_LIST */ meltfrout->tabval[4]))));;

    {
      MELT_LOCATION ("warmelt-normal.melt:6311:/ locexp");
      meltgc_append_list ((melt_ptr_t) ( /*_.RESBINDS__V150*/ meltfptr[149]),
			  (melt_ptr_t) ( /*_.CBIND__V146*/ meltfptr[145]));
    }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6312:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L71*/ meltfnum[50] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6312:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L71*/ meltfnum[50])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L72*/ meltfnum[49] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6312:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[9];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L72*/ meltfnum[49];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6312;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_letrec result clocc=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CLOCC__V148*/ meltfptr[147];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = " nletrec=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.NLETREC__V70*/ meltfptr[28];
	      /*^apply.arg */
	      argtab[7].meltbp_cstring = " resbinds=";
	      /*^apply.arg */
	      argtab[8].meltbp_aptr =
		(melt_ptr_t *) & /*_.RESBINDS__V150*/ meltfptr[149];
	      /*_.MELT_DEBUG_FUN__V152*/ meltfptr[151] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[3])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V151*/ meltfptr[150] =
	      /*_.MELT_DEBUG_FUN__V152*/ meltfptr[151];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6312:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L72*/ meltfnum[49] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V152*/ meltfptr[151] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V151*/ meltfptr[150] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6312:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L71*/ meltfnum[50] = 0;
      /*^clear */
	     /*clear *//*_._IF___V151*/ meltfptr[150] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:6314:/ locexp");

#if MELT_HAVE_DEBUG
      if (melt_need_debug (0))
	melt_dbgshortbacktrace (("normexp_letrec ending"), (15));
#endif
      ;
    }
    ;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6315:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.CLOCC__V148*/ meltfptr[147];;
    MELT_LOCATION ("warmelt-normal.melt:6315:/ putxtraresult");
    if (!meltxrestab_ || !meltxresdescr_)
      goto meltlabend_rout;
    if (meltxresdescr_[0] != MELTBPAR_PTR)
      goto meltlabend_rout;
    if (meltxrestab_[0].meltbp_aptr)
      *(meltxrestab_[0].meltbp_aptr) =
	(melt_ptr_t) ( /*_.RESBINDS__V150*/ meltfptr[149]);
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*_.LET___V141*/ meltfptr[94] = /*_.RETURN___V153*/ meltfptr[151];;

    MELT_LOCATION ("warmelt-normal.melt:6295:/ clear");
	   /*clear *//*_.NLASTBODY__V142*/ meltfptr[141] = 0;
    /*^clear */
	   /*clear *//*_.GET_CTYPE__V143*/ meltfptr[142] = 0;
    /*^clear */
	   /*clear *//*_.NLASTYP__V144*/ meltfptr[143] = 0;
    /*^clear */
	   /*clear *//*_.CSYM__V145*/ meltfptr[144] = 0;
    /*^clear */
	   /*clear *//*_.CBIND__V146*/ meltfptr[145] = 0;
    /*^clear */
	   /*clear *//*_.CLOCC__V148*/ meltfptr[147] = 0;
    /*^clear */
	   /*clear *//*_.RESBINDS__V150*/ meltfptr[149] = 0;
    /*^clear */
	   /*clear *//*_.RETURN___V153*/ meltfptr[151] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*_.MULTI___V122*/ meltfptr[32] = /*_.LET___V141*/ meltfptr[94];;

    MELT_LOCATION ("warmelt-normal.melt:6275:/ clear");
	   /*clear *//*_.IFCPP___V127*/ meltfptr[89] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V131*/ meltfptr[30] = 0;
    /*^clear */
	   /*clear *//*_.LET___V141*/ meltfptr[94] = 0;

    /*^clear */
	   /*clear *//*_.NBODBINDINGS__V124*/ meltfptr[86] = 0;
    /*_.LET___V66*/ meltfptr[51] = /*_.MULTI___V122*/ meltfptr[32];;

    MELT_LOCATION ("warmelt-normal.melt:6192:/ clear");
	   /*clear *//*_#NBCONSBIND__L33*/ meltfnum[10] = 0;
    /*^clear */
	   /*clear *//*_#CURCIX__L34*/ meltfnum[20] = 0;
    /*^clear */
	   /*clear *//*_.NLOCSYMS__V67*/ meltfptr[45] = 0;
    /*^clear */
	   /*clear *//*_.NCBINDTUP__V68*/ meltfptr[40] = 0;
    /*^clear */
	   /*clear *//*_.RECFILLBINDINGS__V69*/ meltfptr[42] = 0;
    /*^clear */
	   /*clear *//*_.NLETREC__V70*/ meltfptr[28] = 0;
    /*^clear */
	   /*clear *//*_.MULTI___V122*/ meltfptr[32] = 0;
    /*_.LET___V14*/ meltfptr[12] = /*_.LET___V66*/ meltfptr[51];;

    MELT_LOCATION ("warmelt-normal.melt:6132:/ clear");
	   /*clear *//*_.SLOC__V15*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.SBINDS__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.SBODY__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_#NBBIND__L6*/ meltfnum[4] = 0;
    /*^clear */
	   /*clear *//*_.BINDLIST__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_.CONSBINDLIST__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_.NEWENV__V20*/ meltfptr[19] = 0;
    /*^clear */
	   /*clear *//*_#stI__L7*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_#MAPLEN__L8*/ meltfnum[7] = 0;
    /*^clear */
	   /*clear *//*_.SYMBEXPRMAP__V21*/ meltfptr[20] = 0;
    /*^clear */
	   /*clear *//*_.SYMBINDMAP__V22*/ meltfptr[21] = 0;
    /*^clear */
	   /*clear *//*_.LET___V66*/ meltfptr[51] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6127:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V14*/ meltfptr[12];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6127:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V6*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[8] = 0;
    /*^clear */
	   /*clear *//*_.LET___V14*/ meltfptr[12] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMEXP_LETREC", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_114_WARMELTmiNORMAL_NORMEXP_LETREC */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un * meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 27
    melt_ptr_t mcfr_varptr[27];
#define MELTFRAM_NBVARNUM 12
    long mcfr_varnum[12];
/*others*/
    const char *loc_CSTRING__o0;
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF_st *)
	meltfirstargp_;
      /* use arguments meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 27; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF nbval 27*/
  meltfram__.mcfr_nbvar = 27 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMAL_PREDEF", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6324:/ getarg");
 /*_.PRED__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V4*/ meltfptr[3])) !=
	      NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_CSTRING)
    goto meltlab_endgetargs;
 /*_?*/ meltfram__.loc_CSTRING__o0 = meltxargtab_[2].meltbp_cstring;
  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6325:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6325:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6325:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6325;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normal_predef pred=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.PRED__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V5*/ meltfptr[4] =
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6325:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V5*/ meltfptr[4] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6325:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V5*/ meltfptr[4] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6326:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6326:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V8*/ meltfptr[4] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6326:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6326) ? (6326) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V8*/ meltfptr[4] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V7*/ meltfptr[5] = /*_._IFELSE___V8*/ meltfptr[4];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6326:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V8*/ meltfptr[4] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V7*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6327:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 7, "NCTX_PREDEFMAP");
  /*_.PREDEFMAP__V10*/ meltfptr[9] = slot;
    };
    ;
 /*_.MAPOBJECT_GET__V11*/ meltfptr[10] =
      /*mapobject_get */
      melt_get_mapobjects ((meltmapobjects_ptr_t)
			   ( /*_.PREDEFMAP__V10*/ meltfptr[9]),
			   (meltobject_ptr_t) ( /*_.PRED__V2*/ meltfptr[1]));;
    MELT_LOCATION ("warmelt-normal.melt:6328:/ cond");
    /*cond */ if ( /*_.MAPOBJECT_GET__V11*/ meltfptr[10])	/*then */
      {
	/*^cond.then */
	/*_.BRK__V12*/ meltfptr[11] = /*_.MAPOBJECT_GET__V11*/ meltfptr[10];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6328:/ cond.else");

	/*^block */
	/*anyblock */
	{

   /*_.GET_RAW_SYMBOL__V13*/ meltfptr[12] =
	    melthookproc_HOOK_NAMED_SYMBOL ((const char *) ( /*_?*/
							    meltfram__.
							    loc_CSTRING__o0),
					    (long) MELT_GET);;
	  /*^compute */
	  /*_.BRK__V12*/ meltfptr[11] =
	    /*_.GET_RAW_SYMBOL__V13*/ meltfptr[12];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6328:/ clear");
	     /*clear *//*_.GET_RAW_SYMBOL__V13*/ meltfptr[12] = 0;
	}
	;
      }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6331:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L4*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6331:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L5*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6331:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L5*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6331;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normal_predef brk=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.BRK__V12*/ meltfptr[11];
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V14*/ meltfptr[12] =
	      /*_.MELT_DEBUG_FUN__V15*/ meltfptr[14];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6331:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L5*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V15*/ meltfptr[14] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V14*/ meltfptr[12] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6331:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V14*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#IS_INTEGERBOX__L6*/ meltfnum[1] =
      (melt_magic_discr ((melt_ptr_t) ( /*_.BRK__V12*/ meltfptr[11])) ==
       MELTOBMAG_INT);;
    MELT_LOCATION ("warmelt-normal.melt:6332:/ cond");
    /*cond */ if ( /*_#IS_INTEGERBOX__L6*/ meltfnum[1])	/*then */
      {
	/*^cond.then */
	/*_#OR___L7*/ meltfnum[0] = /*_#IS_INTEGERBOX__L6*/ meltfnum[1];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6332:/ cond.else");

	/*^block */
	/*anyblock */
	{

   /*_#IS_A__L8*/ meltfnum[7] =
	    melt_is_instance_of ((melt_ptr_t) ( /*_.BRK__V12*/ meltfptr[11]),
				 (melt_ptr_t) (( /*!CLASS_SYMBOL */
						meltfrout->tabval[2])));;
	  /*^compute */
	  /*_#OR___L7*/ meltfnum[0] = /*_#IS_A__L8*/ meltfnum[7];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6332:/ clear");
	     /*clear *//*_#IS_A__L8*/ meltfnum[7] = 0;
	}
	;
      }
    ;
    /*^cond */
    /*cond */ if ( /*_#OR___L7*/ meltfnum[0])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:6333:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  /*^quasiblock */


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_PREDEF */
						   meltfrout->tabval[3])),
				    (2), "CLASS_NREP_PREDEF");
    /*_.INST__V19*/ meltfptr[18] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V19*/ meltfptr[18]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V19*/ meltfptr[18]), (0),
				( /*_.SLOC__V4*/ meltfptr[3]), "NREP_LOC");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NRPREDEF",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V19*/ meltfptr[18]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V19*/ meltfptr[18]), (1),
				( /*_.BRK__V12*/ meltfptr[11]), "NRPREDEF");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V19*/ meltfptr[18],
					"newly made instance");
	  ;
	  /*_.RES__V18*/ meltfptr[17] = /*_.INST__V19*/ meltfptr[18];;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6337:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L9*/ meltfnum[7] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6337:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L9*/ meltfnum[7])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L10*/ meltfnum[9] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6337:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L10*/ meltfnum[9];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6337;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normal_predef res=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.RES__V18*/ meltfptr[17];
		    /*_.MELT_DEBUG_FUN__V21*/ meltfptr[20] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V20*/ meltfptr[19] =
		    /*_.MELT_DEBUG_FUN__V21*/ meltfptr[20];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6337:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L10*/ meltfnum[9] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V21*/ meltfptr[20] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V20*/ meltfptr[19] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6337:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L9*/ meltfnum[7] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V20*/ meltfptr[19] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6338:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.RES__V18*/ meltfptr[17];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6338:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  /*_.LET___V17*/ meltfptr[12] = /*_.RETURN___V22*/ meltfptr[20];;

	  MELT_LOCATION ("warmelt-normal.melt:6333:/ clear");
	     /*clear *//*_.RES__V18*/ meltfptr[17] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V22*/ meltfptr[20] = 0;
	  /*_._IFELSE___V16*/ meltfptr[14] = /*_.LET___V17*/ meltfptr[12];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6332:/ clear");
	     /*clear *//*_.LET___V17*/ meltfptr[12] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

	/*^block */
	/*anyblock */
	{


#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6340:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L11*/ meltfnum[9] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6340:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L11*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L12*/ meltfnum[7] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6340:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[7];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L12*/ meltfnum[7];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6340;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normalpredef fail predefmap=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.PREDEFMAP__V10*/ meltfptr[9];
		    /*^apply.arg */
		    argtab[5].meltbp_cstring = " pred=";
		    /*^apply.arg */
		    argtab[6].meltbp_aptr =
		      (melt_ptr_t *) & /*_.PRED__V2*/ meltfptr[1];
		    /*_.MELT_DEBUG_FUN__V24*/ meltfptr[17] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V23*/ meltfptr[19] =
		    /*_.MELT_DEBUG_FUN__V24*/ meltfptr[17];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6340:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L12*/ meltfnum[7] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V24*/ meltfptr[17] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V23*/ meltfptr[19] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6340:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L11*/ meltfnum[9] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V23*/ meltfptr[19] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;
   /*_.MAKE_STRINGCONST__V25*/ meltfptr[20] =
	    (meltgc_new_stringdup
	     ((meltobject_ptr_t)
	      (( /*!DISCR_VERBATIM_STRING */ meltfrout->tabval[4])),
	      ( /*_?*/ meltfram__.loc_CSTRING__o0)));;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6341:/ locexp");
	    melt_error_str ((melt_ptr_t) ( /*_.SLOC__V4*/ meltfptr[3]),
			    ("not a predef:"),
			    (melt_ptr_t) ( /*_.MAKE_STRINGCONST__V25*/
					  meltfptr[20]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6342:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = ( /*nil */ NULL);;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6342:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  MELT_LOCATION ("warmelt-normal.melt:6339:/ quasiblock");


	  /*_.PROGN___V27*/ meltfptr[17] = /*_.RETURN___V26*/ meltfptr[12];;
	  /*^compute */
	  /*_._IFELSE___V16*/ meltfptr[14] = /*_.PROGN___V27*/ meltfptr[17];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6332:/ clear");
	     /*clear *//*_.MAKE_STRINGCONST__V25*/ meltfptr[20] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V26*/ meltfptr[12] = 0;
	  /*^clear */
	     /*clear *//*_.PROGN___V27*/ meltfptr[17] = 0;
	}
	;
      }
    ;
    /*_.LET___V9*/ meltfptr[4] = /*_._IFELSE___V16*/ meltfptr[14];;

    MELT_LOCATION ("warmelt-normal.melt:6327:/ clear");
	   /*clear *//*_.PREDEFMAP__V10*/ meltfptr[9] = 0;
    /*^clear */
	   /*clear *//*_.MAPOBJECT_GET__V11*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.BRK__V12*/ meltfptr[11] = 0;
    /*^clear */
	   /*clear *//*_#IS_INTEGERBOX__L6*/ meltfnum[1] = 0;
    /*^clear */
	   /*clear *//*_#OR___L7*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_._IFELSE___V16*/ meltfptr[14] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6324:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V9*/ meltfptr[4];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6324:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V7*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.LET___V9*/ meltfptr[4] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMAL_PREDEF", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_115_WARMELTmiNORMAL_NORMAL_PREDEF */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA (meltclosure_ptr_t meltclosp_,
						 melt_ptr_t meltfirstargp_,
						 const melt_argdescr_cell_t
						 meltxargdescr_[],
						 union meltparam_un *
						 meltxargtab_,
						 const melt_argdescr_cell_t
						 meltxresdescr_[],
						 union meltparam_un *
						 meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 36
    melt_ptr_t mcfr_varptr[36];
#define MELTFRAM_NBVARNUM 11
    long mcfr_varnum[11];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA_st
	 *) meltfirstargp_;
      /* use arguments meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 36; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA nbval 36*/
  meltfram__.mcfr_nbvar = 36 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMAL_SYMBOL_DATA", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6347:/ getarg");
 /*_.SYM__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V4*/ meltfptr[3])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6349:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6349:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6349:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6349;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normal_symbol_data start sym=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.SYM__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V5*/ meltfptr[4] =
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6349:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V5*/ meltfptr[4] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6349:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V5*/ meltfptr[4] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6350:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_.DISCRIM__V8*/ meltfptr[4] =
	((melt_ptr_t)
	 (melt_discr ((melt_ptr_t) ( /*_.SYM__V2*/ meltfptr[1]))));;
      /*^compute */
   /*_#eqeq__L3*/ meltfnum[1] =
	(( /*_.DISCRIM__V8*/ meltfptr[4]) ==
	 (( /*!CLASS_SYMBOL */ meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6350:/ cond");
      /*cond */ if ( /*_#eqeq__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6350:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check symb"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6350) ? (6350) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V7*/ meltfptr[5] = /*_._IFELSE___V9*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6350:/ clear");
	     /*clear *//*_.DISCRIM__V8*/ meltfptr[4] = 0;
      /*^clear */
	     /*clear *//*_#eqeq__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V7*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6351:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:6351:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6351:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6351) ? (6351) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[4] = /*_._IFELSE___V11*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6351:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[4] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6352:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 8, "NCTX_VALMAP");
  /*_.VALMAP__V13*/ meltfptr[12] = slot;
    };
    ;
 /*_.OSYDATA__V14*/ meltfptr[13] =
      /*mapobject_get */
      melt_get_mapobjects ((meltmapobjects_ptr_t)
			   ( /*_.VALMAP__V13*/ meltfptr[12]),
			   (meltobject_ptr_t) ( /*_.SYM__V2*/ meltfptr[1]));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6354:/ cond");
    /*cond */ if ( /*_.OSYDATA__V14*/ meltfptr[13])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{


#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6356:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L5*/ meltfnum[1] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6356:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L5*/ meltfnum[1])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L6*/ meltfnum[0] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6356:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L6*/ meltfnum[0];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6356;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normal_symbol_data found osydata";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.OSYDATA__V14*/ meltfptr[13];
		    /*_.MELT_DEBUG_FUN__V17*/ meltfptr[16] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V16*/ meltfptr[15] =
		    /*_.MELT_DEBUG_FUN__V17*/ meltfptr[16];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6356:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L6*/ meltfnum[0] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V17*/ meltfptr[16] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V16*/ meltfptr[15] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6356:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L5*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V16*/ meltfptr[15] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6357:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.OSYDATA__V14*/ meltfptr[13];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6357:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  MELT_LOCATION ("warmelt-normal.melt:6355:/ quasiblock");


	  /*_.PROGN___V19*/ meltfptr[15] = /*_.RETURN___V18*/ meltfptr[16];;
	  /*^compute */
	  /*_._IFELSE___V15*/ meltfptr[14] = /*_.PROGN___V19*/ meltfptr[15];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6354:/ clear");
	     /*clear *//*_.RETURN___V18*/ meltfptr[16] = 0;
	  /*^clear */
	     /*clear *//*_.PROGN___V19*/ meltfptr[15] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:6358:/ quasiblock");


   /*_#SYHASH__L7*/ meltfnum[0] =
	    (melt_obj_hash ((melt_ptr_t) ( /*_.SYM__V2*/ meltfptr[1])));;
	  MELT_LOCATION ("warmelt-normal.melt:6359:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.SYM__V2*/ meltfptr[1]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 1, "NAMED_NAME");
    /*_.SYNAMSTR__V21*/ meltfptr[15] = slot;
	  };
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6363:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V3*/ meltfptr[2];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[2].meltbp_cstring = "discr_string";
	    /*_.NORMAL_PREDEF__V22*/ meltfptr[21] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_PREDEF */ meltfrout->tabval[4])),
			  (melt_ptr_t) (( /*!DISCR_STRING */ meltfrout->
					 tabval[5])),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			   ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6362:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_DATASTRING */ meltfrout->tabval[3])), (6), "CLASS_NREP_DATASTRING");
    /*_.INST__V24*/ meltfptr[23] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V24*/ meltfptr[23]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V24*/ meltfptr[23]), (4),
				( /*_.NORMAL_PREDEF__V22*/ meltfptr[21]),
				"NDATA_DISCRX");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NSTR_STRING",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V24*/ meltfptr[23]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V24*/ meltfptr[23]), (5),
				( /*_.SYNAMSTR__V21*/ meltfptr[15]),
				"NSTR_STRING");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V24*/ meltfptr[23],
					"newly made instance");
	  ;
	  /*_.SYNAMSTRDATA__V23*/ meltfptr[22] =
	    /*_.INST__V24*/ meltfptr[23];;
	  MELT_LOCATION ("warmelt-normal.melt:6367:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj =
	      (melt_ptr_t) (( /*!CLASS_SYMBOL */ meltfrout->
			     tabval[1])) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 6, "CLASS_FIELDS");
    /*_.CLASS_FIELDS__V25*/ meltfptr[24] = slot;
	  };
	  ;
   /*_#MULTIPLE_LENGTH__L8*/ meltfnum[1] =
	    (melt_multiple_length
	     ((melt_ptr_t) ( /*_.CLASS_FIELDS__V25*/ meltfptr[24])));;
	  /*^compute */
   /*_.SYSLOTS__V26*/ meltfptr[25] =
	    (meltgc_new_multiple
	     ((meltobject_ptr_t)
	      (( /*!DISCR_MULTIPLE */ meltfrout->tabval[6])),
	      ( /*_#MULTIPLE_LENGTH__L8*/ meltfnum[1])));;

	  MELT_CHECK_SIGNAL ();
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6370:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V3*/ meltfptr[2];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[2].meltbp_cstring = "class_symbol";
	    /*_.NORMAL_PREDEF__V27*/ meltfptr[26] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_PREDEF */ meltfrout->tabval[4])),
			  (melt_ptr_t) (( /*!CLASS_SYMBOL */ meltfrout->
					 tabval[1])),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			   ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
   /*_.MAKE_INTEGERBOX__V28*/ meltfptr[27] =
	    (meltgc_new_int
	     ((meltobject_ptr_t)
	      (( /*!DISCR_INTEGER */ meltfrout->tabval[8])),
	      ( /*_#SYHASH__L7*/ meltfnum[0])));;
	  MELT_LOCATION ("warmelt-normal.melt:6368:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_DATASYMBOL */ meltfrout->tabval[7])), (10), "CLASS_NREP_DATASYMBOL");
    /*_.INST__V30*/ meltfptr[29] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_NAME",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V30*/ meltfptr[29]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V30*/ meltfptr[29]), (1),
				( /*_.SYM__V2*/ meltfptr[1]), "NDATA_NAME");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V30*/ meltfptr[29]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V30*/ meltfptr[29]), (4),
				( /*_.NORMAL_PREDEF__V27*/ meltfptr[26]),
				"NDATA_DISCRX");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NINST_HASH",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V30*/ meltfptr[29]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V30*/ meltfptr[29]), (7),
				( /*_.MAKE_INTEGERBOX__V28*/ meltfptr[27]),
				"NINST_HASH");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NINST_SLOTS",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V30*/ meltfptr[29]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V30*/ meltfptr[29]), (8),
				( /*_.SYSLOTS__V26*/ meltfptr[25]),
				"NINST_SLOTS");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDSY_NAMESTR",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V30*/ meltfptr[29]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V30*/ meltfptr[29]), (9),
				( /*_.SYNAMSTR__V21*/ meltfptr[15]),
				"NDSY_NAMESTR");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V30*/ meltfptr[29],
					"newly made instance");
	  ;
	  /*_.SYDATA__V29*/ meltfptr[28] = /*_.INST__V30*/ meltfptr[29];;
	  /*^compute */
   /*_#OBJ_NUM__L9*/ meltfnum[8] =
	    (melt_obj_num
	     ((melt_ptr_t) (( /*!NAMED_NAME */ meltfrout->tabval[9]))));;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6375:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.SYSLOTS__V26*/ meltfptr[25]),
				     ( /*_#OBJ_NUM__L9*/ meltfnum[8]),
				     (melt_ptr_t) ( /*_.SYNAMSTRDATA__V23*/
						   meltfptr[22]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6376:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SYDATA__V29*/ meltfptr[28];
	    /*_.ADD_NCTX_DATA__V31*/ meltfptr[30] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!ADD_NCTX_DATA */ meltfrout->tabval[10])),
			  (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6377:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SYNAMSTRDATA__V23*/ meltfptr[22];
	    /*_.ADD_NCTX_DATA__V32*/ meltfptr[31] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!ADD_NCTX_DATA */ meltfrout->tabval[10])),
			  (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6378:/ locexp");
	    meltgc_put_mapobjects ((meltmapobjects_ptr_t)
				   ( /*_.VALMAP__V13*/ meltfptr[12]),
				   (meltobject_ptr_t) ( /*_.SYM__V2*/
						       meltfptr[1]),
				   (melt_ptr_t) ( /*_.SYDATA__V29*/
						 meltfptr[28]));
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6379:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 4, "NCTX_SYMBMAP");
    /*_.NCTX_SYMBMAP__V33*/ meltfptr[32] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    meltgc_put_mapstrings ((struct meltmapstrings_st *) ( /*_.NCTX_SYMBMAP__V33*/ meltfptr[32]),
				   melt_string_str ((melt_ptr_t)
						    ( /*_.SYNAMSTR__V21*/
						     meltfptr[15])),
				   (melt_ptr_t) ( /*_.SYDATA__V29*/
						 meltfptr[28]));
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6380:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L10*/ meltfnum[9] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6380:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[9])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L11*/ meltfnum[10] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6380:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L11*/ meltfnum[10];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6380;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normal_symbol_data return sydata";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.SYDATA__V29*/ meltfptr[28];
		    /*_.MELT_DEBUG_FUN__V35*/ meltfptr[34] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V34*/ meltfptr[33] =
		    /*_.MELT_DEBUG_FUN__V35*/ meltfptr[34];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6380:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[10] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V35*/ meltfptr[34] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V34*/ meltfptr[33] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6380:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[9] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V34*/ meltfptr[33] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6381:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.SYDATA__V29*/ meltfptr[28];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6381:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  /*_.LET___V20*/ meltfptr[16] = /*_.RETURN___V36*/ meltfptr[34];;

	  MELT_LOCATION ("warmelt-normal.melt:6358:/ clear");
	     /*clear *//*_#SYHASH__L7*/ meltfnum[0] = 0;
	  /*^clear */
	     /*clear *//*_.SYNAMSTR__V21*/ meltfptr[15] = 0;
	  /*^clear */
	     /*clear *//*_.NORMAL_PREDEF__V22*/ meltfptr[21] = 0;
	  /*^clear */
	     /*clear *//*_.SYNAMSTRDATA__V23*/ meltfptr[22] = 0;
	  /*^clear */
	     /*clear *//*_.CLASS_FIELDS__V25*/ meltfptr[24] = 0;
	  /*^clear */
	     /*clear *//*_#MULTIPLE_LENGTH__L8*/ meltfnum[1] = 0;
	  /*^clear */
	     /*clear *//*_.SYSLOTS__V26*/ meltfptr[25] = 0;
	  /*^clear */
	     /*clear *//*_.NORMAL_PREDEF__V27*/ meltfptr[26] = 0;
	  /*^clear */
	     /*clear *//*_.MAKE_INTEGERBOX__V28*/ meltfptr[27] = 0;
	  /*^clear */
	     /*clear *//*_.SYDATA__V29*/ meltfptr[28] = 0;
	  /*^clear */
	     /*clear *//*_#OBJ_NUM__L9*/ meltfnum[8] = 0;
	  /*^clear */
	     /*clear *//*_.ADD_NCTX_DATA__V31*/ meltfptr[30] = 0;
	  /*^clear */
	     /*clear *//*_.ADD_NCTX_DATA__V32*/ meltfptr[31] = 0;
	  /*^clear */
	     /*clear *//*_.NCTX_SYMBMAP__V33*/ meltfptr[32] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V36*/ meltfptr[34] = 0;
	  /*_._IFELSE___V15*/ meltfptr[14] = /*_.LET___V20*/ meltfptr[16];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6354:/ clear");
	     /*clear *//*_.LET___V20*/ meltfptr[16] = 0;
	}
	;
      }
    ;
    /*_.LET___V12*/ meltfptr[8] = /*_._IFELSE___V15*/ meltfptr[14];;

    MELT_LOCATION ("warmelt-normal.melt:6352:/ clear");
	   /*clear *//*_.VALMAP__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.OSYDATA__V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_._IFELSE___V15*/ meltfptr[14] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6347:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V12*/ meltfptr[8];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6347:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V7*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[4] = 0;
    /*^clear */
	   /*clear *//*_.LET___V12*/ meltfptr[8] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMAL_SYMBOL_DATA", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_116_WARMELTmiNORMAL_NORMAL_SYMBOL_DATA */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA (meltclosure_ptr_t
						  meltclosp_,
						  melt_ptr_t meltfirstargp_,
						  const melt_argdescr_cell_t
						  meltxargdescr_[],
						  union meltparam_un *
						  meltxargtab_,
						  const melt_argdescr_cell_t
						  meltxresdescr_[],
						  union meltparam_un *
						  meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 34
    melt_ptr_t mcfr_varptr[34];
#define MELTFRAM_NBVARNUM 11
    long mcfr_varnum[11];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA_st
	 *) meltfirstargp_;
      /* use arguments meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 34; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA nbval 34*/
  meltfram__.mcfr_nbvar = 34 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMAL_KEYWORD_DATA", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6387:/ getarg");
 /*_.KEYW__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.SLOC__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.SLOC__V4*/ meltfptr[3])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6389:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6389:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6389:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6389;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normal_keyword_data keyw=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.KEYW__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V5*/ meltfptr[4] =
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6389:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V5*/ meltfptr[4] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6389:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V5*/ meltfptr[4] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6390:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_.DISCRIM__V8*/ meltfptr[4] =
	((melt_ptr_t)
	 (melt_discr ((melt_ptr_t) ( /*_.KEYW__V2*/ meltfptr[1]))));;
      /*^compute */
   /*_#eqeq__L3*/ meltfnum[1] =
	(( /*_.DISCRIM__V8*/ meltfptr[4]) ==
	 (( /*!CLASS_KEYWORD */ meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6390:/ cond");
      /*cond */ if ( /*_#eqeq__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6390:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check keywb"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6390) ? (6390) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V7*/ meltfptr[5] = /*_._IFELSE___V9*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6390:/ clear");
	     /*clear *//*_.DISCRIM__V8*/ meltfptr[4] = 0;
      /*^clear */
	     /*clear *//*_#eqeq__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V7*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6391:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:6391:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[8] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6391:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6391) ? (6391) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[8] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[4] = /*_._IFELSE___V11*/ meltfptr[8];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6391:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[8] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[4] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6392:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 8, "NCTX_VALMAP");
  /*_.VALMAP__V13*/ meltfptr[12] = slot;
    };
    ;
 /*_.OSYDATA__V14*/ meltfptr[13] =
      /*mapobject_get */
      melt_get_mapobjects ((meltmapobjects_ptr_t)
			   ( /*_.VALMAP__V13*/ meltfptr[12]),
			   (meltobject_ptr_t) ( /*_.KEYW__V2*/ meltfptr[1]));;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6394:/ cond");
    /*cond */ if ( /*_.OSYDATA__V14*/ meltfptr[13])	/*then */
      {
	/*^cond.then */
	/*_._IFELSE___V15*/ meltfptr[14] = /*_.OSYDATA__V14*/ meltfptr[13];;
      }
    else
      {
	MELT_LOCATION ("warmelt-normal.melt:6394:/ cond.else");

	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:6395:/ quasiblock");


   /*_#SYHASH__L5*/ meltfnum[1] =
	    (melt_obj_hash ((melt_ptr_t) ( /*_.KEYW__V2*/ meltfptr[1])));;
	  MELT_LOCATION ("warmelt-normal.melt:6396:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.KEYW__V2*/ meltfptr[1]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 1, "NAMED_NAME");
    /*_.SYNAMSTR__V17*/ meltfptr[16] = slot;
	  };
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6400:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V3*/ meltfptr[2];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[2].meltbp_cstring = "discr_string";
	    /*_.NORMAL_PREDEF__V18*/ meltfptr[17] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_PREDEF */ meltfrout->tabval[4])),
			  (melt_ptr_t) (( /*!DISCR_STRING */ meltfrout->
					 tabval[5])),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			   ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6399:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_DATASTRING */ meltfrout->tabval[3])), (6), "CLASS_NREP_DATASTRING");
    /*_.INST__V20*/ meltfptr[19] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V20*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (4),
				( /*_.NORMAL_PREDEF__V18*/ meltfptr[17]),
				"NDATA_DISCRX");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NSTR_STRING",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V20*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V20*/ meltfptr[19]), (5),
				( /*_.SYNAMSTR__V17*/ meltfptr[16]),
				"NSTR_STRING");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V20*/ meltfptr[19],
					"newly made instance");
	  ;
	  /*_.SYNAMSTRDATA__V19*/ meltfptr[18] =
	    /*_.INST__V20*/ meltfptr[19];;
	  MELT_LOCATION ("warmelt-normal.melt:6404:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj =
	      (melt_ptr_t) (( /*!CLASS_KEYWORD */ meltfrout->
			     tabval[1])) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 6, "CLASS_FIELDS");
    /*_.CLASS_FIELDS__V21*/ meltfptr[20] = slot;
	  };
	  ;
   /*_#MULTIPLE_LENGTH__L6*/ meltfnum[0] =
	    (melt_multiple_length
	     ((melt_ptr_t) ( /*_.CLASS_FIELDS__V21*/ meltfptr[20])));;
	  /*^compute */
   /*_.SYSLOTS__V22*/ meltfptr[21] =
	    (meltgc_new_multiple
	     ((meltobject_ptr_t)
	      (( /*!DISCR_MULTIPLE */ meltfrout->tabval[6])),
	      ( /*_#MULTIPLE_LENGTH__L6*/ meltfnum[0])));;

	  MELT_CHECK_SIGNAL ();
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6407:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[3];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V3*/ meltfptr[2];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SLOC__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[2].meltbp_cstring = "class_keyword";
	    /*_.NORMAL_PREDEF__V23*/ meltfptr[22] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_PREDEF */ meltfrout->tabval[4])),
			  (melt_ptr_t) (( /*!CLASS_KEYWORD */ meltfrout->
					 tabval[1])),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			   ""), argtab, "", (union meltparam_un *) 0);
	  }
	  ;
   /*_.MAKE_INTEGERBOX__V24*/ meltfptr[23] =
	    (meltgc_new_int
	     ((meltobject_ptr_t)
	      (( /*!DISCR_INTEGER */ meltfrout->tabval[8])),
	      ( /*_#SYHASH__L5*/ meltfnum[1])));;
	  MELT_LOCATION ("warmelt-normal.melt:6405:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_DATAKEYWORD */ meltfrout->tabval[7])), (10), "CLASS_NREP_DATAKEYWORD");
    /*_.INST__V26*/ meltfptr[25] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_NAME",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V26*/ meltfptr[25]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V26*/ meltfptr[25]), (1),
				( /*_.KEYW__V2*/ meltfptr[1]), "NDATA_NAME");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V26*/ meltfptr[25]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V26*/ meltfptr[25]), (4),
				( /*_.NORMAL_PREDEF__V23*/ meltfptr[22]),
				"NDATA_DISCRX");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NINST_HASH",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V26*/ meltfptr[25]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V26*/ meltfptr[25]), (7),
				( /*_.MAKE_INTEGERBOX__V24*/ meltfptr[23]),
				"NINST_HASH");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NINST_SLOTS",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V26*/ meltfptr[25]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V26*/ meltfptr[25]), (8),
				( /*_.SYSLOTS__V22*/ meltfptr[21]),
				"NINST_SLOTS");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NDSY_NAMESTR",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V26*/ meltfptr[25]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V26*/ meltfptr[25]), (9),
				( /*_.SYNAMSTR__V17*/ meltfptr[16]),
				"NDSY_NAMESTR");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V26*/ meltfptr[25],
					"newly made instance");
	  ;
	  /*_.SYDATA__V25*/ meltfptr[24] = /*_.INST__V26*/ meltfptr[25];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6412:/ locexp");

#if MELT_HAVE_DEBUG
	    if (melt_need_debug (0))
	      melt_dbgshortbacktrace (("normal_keyword_data"), (8));
#endif
	    ;
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6413:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L7*/ meltfnum[6] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6413:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L7*/ meltfnum[6])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L8*/ meltfnum[7] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6413:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L8*/ meltfnum[7];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6413;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normal_keyword_data sydata=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.SYDATA__V25*/ meltfptr[24];
		    /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V27*/ meltfptr[26] =
		    /*_.MELT_DEBUG_FUN__V28*/ meltfptr[27];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6413:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L8*/ meltfnum[7] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V28*/ meltfptr[27] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V27*/ meltfptr[26] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6413:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L7*/ meltfnum[6] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V27*/ meltfptr[26] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;
   /*_#OBJ_NUM__L9*/ meltfnum[7] =
	    (melt_obj_num
	     ((melt_ptr_t) (( /*!NAMED_NAME */ meltfrout->tabval[9]))));;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6414:/ locexp");
	    meltgc_multiple_put_nth ((melt_ptr_t)
				     ( /*_.SYSLOTS__V22*/ meltfptr[21]),
				     ( /*_#OBJ_NUM__L9*/ meltfnum[7]),
				     (melt_ptr_t) ( /*_.SYNAMSTRDATA__V19*/
						   meltfptr[18]));
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6415:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SYDATA__V25*/ meltfptr[24];
	    /*_.ADD_NCTX_DATA__V29*/ meltfptr[27] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!ADD_NCTX_DATA */ meltfrout->tabval[10])),
			  (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6416:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[1];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.SYNAMSTRDATA__V19*/ meltfptr[18];
	    /*_.ADD_NCTX_DATA__V30*/ meltfptr[26] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!ADD_NCTX_DATA */ meltfrout->tabval[10])),
			  (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]),
			  (MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6417:/ locexp");
	    meltgc_put_mapobjects ((meltmapobjects_ptr_t)
				   ( /*_.VALMAP__V13*/ meltfptr[12]),
				   (meltobject_ptr_t) ( /*_.KEYW__V2*/
						       meltfptr[1]),
				   (melt_ptr_t) ( /*_.SYDATA__V25*/
						 meltfptr[24]));
	  }
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6418:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.NCX__V3*/ meltfptr[2]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 5, "NCTX_KEYWMAP");
    /*_.NCTX_KEYWMAP__V31*/ meltfptr[30] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    meltgc_put_mapstrings ((struct meltmapstrings_st *) ( /*_.NCTX_KEYWMAP__V31*/ meltfptr[30]),
				   melt_string_str ((melt_ptr_t)
						    ( /*_.SYNAMSTR__V17*/
						     meltfptr[16])),
				   (melt_ptr_t) ( /*_.SYDATA__V25*/
						 meltfptr[24]));
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6419:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L10*/ meltfnum[6] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6419:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L10*/ meltfnum[6])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L11*/ meltfnum[10] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6419:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L11*/ meltfnum[10];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6419;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normal_keyword_data return sydata=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.SYDATA__V25*/ meltfptr[24];
		    /*_.MELT_DEBUG_FUN__V33*/ meltfptr[32] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V32*/ meltfptr[31] =
		    /*_.MELT_DEBUG_FUN__V33*/ meltfptr[32];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6419:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L11*/ meltfnum[10] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V33*/ meltfptr[32] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V32*/ meltfptr[31] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6419:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L10*/ meltfnum[6] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V32*/ meltfptr[31] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6420:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.SYDATA__V25*/ meltfptr[24];;

	  {
	    MELT_LOCATION ("warmelt-normal.melt:6420:/ locexp");
	    /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
	    if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	      melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
	    ;
	  }
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  /*_.LET___V16*/ meltfptr[15] = /*_.RETURN___V34*/ meltfptr[32];;

	  MELT_LOCATION ("warmelt-normal.melt:6395:/ clear");
	     /*clear *//*_#SYHASH__L5*/ meltfnum[1] = 0;
	  /*^clear */
	     /*clear *//*_.SYNAMSTR__V17*/ meltfptr[16] = 0;
	  /*^clear */
	     /*clear *//*_.NORMAL_PREDEF__V18*/ meltfptr[17] = 0;
	  /*^clear */
	     /*clear *//*_.SYNAMSTRDATA__V19*/ meltfptr[18] = 0;
	  /*^clear */
	     /*clear *//*_.CLASS_FIELDS__V21*/ meltfptr[20] = 0;
	  /*^clear */
	     /*clear *//*_#MULTIPLE_LENGTH__L6*/ meltfnum[0] = 0;
	  /*^clear */
	     /*clear *//*_.SYSLOTS__V22*/ meltfptr[21] = 0;
	  /*^clear */
	     /*clear *//*_.NORMAL_PREDEF__V23*/ meltfptr[22] = 0;
	  /*^clear */
	     /*clear *//*_.MAKE_INTEGERBOX__V24*/ meltfptr[23] = 0;
	  /*^clear */
	     /*clear *//*_.SYDATA__V25*/ meltfptr[24] = 0;
	  /*^clear */
	     /*clear *//*_#OBJ_NUM__L9*/ meltfnum[7] = 0;
	  /*^clear */
	     /*clear *//*_.ADD_NCTX_DATA__V29*/ meltfptr[27] = 0;
	  /*^clear */
	     /*clear *//*_.ADD_NCTX_DATA__V30*/ meltfptr[26] = 0;
	  /*^clear */
	     /*clear *//*_.NCTX_KEYWMAP__V31*/ meltfptr[30] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V34*/ meltfptr[32] = 0;
	  /*_._IFELSE___V15*/ meltfptr[14] = /*_.LET___V16*/ meltfptr[15];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6394:/ clear");
	     /*clear *//*_.LET___V16*/ meltfptr[15] = 0;
	}
	;
      }
    ;
    /*_.LET___V12*/ meltfptr[8] = /*_._IFELSE___V15*/ meltfptr[14];;

    MELT_LOCATION ("warmelt-normal.melt:6392:/ clear");
	   /*clear *//*_.VALMAP__V13*/ meltfptr[12] = 0;
    /*^clear */
	   /*clear *//*_.OSYDATA__V14*/ meltfptr[13] = 0;
    /*^clear */
	   /*clear *//*_._IFELSE___V15*/ meltfptr[14] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6387:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V12*/ meltfptr[8];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6387:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V7*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[4] = 0;
    /*^clear */
	   /*clear *//*_.LET___V12*/ meltfptr[8] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMAL_KEYWORD_DATA", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_117_WARMELTmiNORMAL_NORMAL_KEYWORD_DATA */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS (meltclosure_ptr_t meltclosp_,
						melt_ptr_t meltfirstargp_,
						const melt_argdescr_cell_t
						meltxargdescr_[],
						union meltparam_un *
						meltxargtab_,
						const melt_argdescr_cell_t
						meltxresdescr_[],
						union meltparam_un *
						meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 11
    melt_ptr_t mcfr_varptr[11];
#define MELTFRAM_NBVARNUM 6
    long mcfr_varnum[6];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS_st *)
	meltfirstargp_;
      /* use arguments meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 11; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS nbval 11*/
  meltfram__.mcfr_nbvar = 11 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("CREATE_DATA_SLOTS", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6425:/ getarg");
 /*_.CLA__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;
  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6426:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L1*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.CLA__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_CLASS */ meltfrout->
					    tabval[0])));;
      MELT_LOCATION ("warmelt-normal.melt:6426:/ cond");
      /*cond */ if ( /*_#IS_A__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V4*/ meltfptr[3] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6426:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check cla"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6426) ? (6426) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V4*/ meltfptr[3] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V3*/ meltfptr[2] = /*_._IFELSE___V4*/ meltfptr[3];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6426:/ clear");
	     /*clear *//*_#IS_A__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V4*/ meltfptr[3] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V3*/ meltfptr[2] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6427:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L2*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6427:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L2*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L3*/ meltfnum[2] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6427:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L3*/ meltfnum[2];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6427;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "create_data_slots cla";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.CLA__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[1])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V5*/ meltfptr[3] =
	      /*_.MELT_DEBUG_FUN__V6*/ meltfptr[5];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6427:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L3*/ meltfnum[2] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V6*/ meltfptr[5] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V5*/ meltfptr[3] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6427:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L2*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V5*/ meltfptr[3] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6428:/ quasiblock");


    MELT_LOCATION ("warmelt-normal.melt:6429:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.CLA__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 6, "CLASS_FIELDS");
  /*_.CLASS_FIELDS__V8*/ meltfptr[3] = slot;
    };
    ;
 /*_#MULTIPLE_LENGTH__L4*/ meltfnum[2] =
      (melt_multiple_length
       ((melt_ptr_t) ( /*_.CLASS_FIELDS__V8*/ meltfptr[3])));;
    /*^compute */
 /*_.TUPSLO__V9*/ meltfptr[8] =
      (meltgc_new_multiple
       ((meltobject_ptr_t) (( /*!DISCR_MULTIPLE */ meltfrout->tabval[2])),
	( /*_#MULTIPLE_LENGTH__L4*/ meltfnum[2])));;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6431:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L5*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6431:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L5*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L6*/ meltfnum[5] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6431:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L6*/ meltfnum[5];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6431;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "create_data_slots tupslo";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.TUPSLO__V9*/ meltfptr[8];
	      /*_.MELT_DEBUG_FUN__V11*/ meltfptr[10] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[1])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V10*/ meltfptr[9] =
	      /*_.MELT_DEBUG_FUN__V11*/ meltfptr[10];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6431:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L6*/ meltfnum[5] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V11*/ meltfptr[10] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V10*/ meltfptr[9] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6431:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L5*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V10*/ meltfptr[9] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;
    /*_.LET___V7*/ meltfptr[5] = /*_.TUPSLO__V9*/ meltfptr[8];;

    MELT_LOCATION ("warmelt-normal.melt:6428:/ clear");
	   /*clear *//*_.CLASS_FIELDS__V8*/ meltfptr[3] = 0;
    /*^clear */
	   /*clear *//*_#MULTIPLE_LENGTH__L4*/ meltfnum[2] = 0;
    /*^clear */
	   /*clear *//*_.TUPSLO__V9*/ meltfptr[8] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6425:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V7*/ meltfptr[5];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6425:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V3*/ meltfptr[2] = 0;
    /*^clear */
	   /*clear *//*_.LET___V7*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("CREATE_DATA_SLOTS", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_118_WARMELTmiNORMAL_CREATE_DATA_SLOTS */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT (meltclosure_ptr_t meltclosp_,
					     melt_ptr_t meltfirstargp_,
					     const melt_argdescr_cell_t
					     meltxargdescr_[],
					     union meltparam_un *
					     meltxargtab_,
					     const melt_argdescr_cell_t
					     meltxresdescr_[],
					     union meltparam_un *
					     meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 9
    melt_ptr_t mcfr_varptr[9];
#define MELTFRAM_NBVARNUM 3
    long mcfr_varnum[3];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT_st *)
	meltfirstargp_;
      /* use arguments meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 9; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT nbval 9*/
  meltfram__.mcfr_nbvar = 9 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("FILL_DATA_SLOT", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6436:/ getarg");
 /*_.DI__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.FIELD__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.FIELD__V3*/ meltfptr[2])) !=
	      NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.VAL__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.VAL__V4*/ meltfptr[3])) != NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6437:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L1*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.DI__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_NREP_DATAINSTANCE */
					    meltfrout->tabval[0])));;
      MELT_LOCATION ("warmelt-normal.melt:6437:/ cond");
      /*cond */ if ( /*_#IS_A__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V6*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6437:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check di"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6437) ? (6437) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V6*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V5*/ meltfptr[4] = /*_._IFELSE___V6*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6437:/ clear");
	     /*clear *//*_#IS_A__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V5*/ meltfptr[4] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6438:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L2*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.FIELD__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_FIELD */ meltfrout->
					    tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6438:/ cond");
      /*cond */ if ( /*_#IS_A__L2*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V8*/ meltfptr[7] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6438:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check field"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6438) ? (6438) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V8*/ meltfptr[7] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V7*/ meltfptr[5] = /*_._IFELSE___V8*/ meltfptr[7];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6438:/ clear");
	     /*clear *//*_#IS_A__L2*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V8*/ meltfptr[7] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V7*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;
    MELT_LOCATION ("warmelt-normal.melt:6439:/ quasiblock");


 /*_#FIX__L3*/ meltfnum[0] =
      (melt_obj_num ((melt_ptr_t) ( /*_.FIELD__V3*/ meltfptr[2])));;
    MELT_LOCATION ("warmelt-normal.melt:6440:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.DI__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 8, "NINST_SLOTS");
  /*_.SLOTS__V9*/ meltfptr[7] = slot;
    };
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:6441:/ locexp");
      meltgc_multiple_put_nth ((melt_ptr_t) ( /*_.SLOTS__V9*/ meltfptr[7]),
			       ( /*_#FIX__L3*/ meltfnum[0]),
			       (melt_ptr_t) ( /*_.VAL__V4*/ meltfptr[3]));
    }
    ;

    MELT_LOCATION ("warmelt-normal.melt:6439:/ clear");
	   /*clear *//*_#FIX__L3*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.SLOTS__V9*/ meltfptr[7] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    /*epilog */

    MELT_LOCATION ("warmelt-normal.melt:6436:/ clear");
	   /*clear *//*_.IFCPP___V5*/ meltfptr[4] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V7*/ meltfptr[5] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("FILL_DATA_SLOT", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*noretval */ NULL);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_119_WARMELTmiNORMAL_FILL_DATA_SLOT */





melt_ptr_t MELT_MODULE_VISIBILITY
meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE (meltclosure_ptr_t meltclosp_,
					    melt_ptr_t meltfirstargp_,
					    const melt_argdescr_cell_t
					    meltxargdescr_[],
					    union meltparam_un * meltxargtab_,
					    const melt_argdescr_cell_t
					    meltxresdescr_[],
					    union meltparam_un * meltxrestab_)
{
  long
    current_blocklevel_signals_meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE_melt
    = melt_blocklevel_signals;

#if MELT_HAVE_DEBUG
  static long melt_call_counter__;
  long melt_thiscallcounter__ ATTRIBUTE_UNUSED = ++melt_call_counter__;
#undef meltcallcount
#define meltcallcount melt_thiscallcounter__
#else
#undef meltcallcount
#define meltcallcount 0L
#endif

  struct meltframe_meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE_st
  {
    int mcfr_nbvar;
#if MELT_HAVE_DEBUG
    const char *mcfr_flocs;
#else				/*!MELT_HAVE_DEBUG */
    const char *mcfr_unusedflocs;
#endif				/*MELT_HAVE_DEBUG */
    struct meltclosure_st *mcfr_clos;
    struct excepth_melt_st *mcfr_exh;
    struct melt_callframe_st *mcfr_prev;
#undef MELTFRAM_NBVARPTR
#undef MELTFRAM_NBVARNUM
#define MELTFRAM_NBVARPTR 49
    melt_ptr_t mcfr_varptr[49];
#define MELTFRAM_NBVARNUM 15
    long mcfr_varnum[15];
/*others*/
    long _spare_;
  }
   *meltframptr_ = 0, meltfram__;	/*declfrastruct */
#define meltframe meltfram__
  /*meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE is not initial declstructinit */
  if (MELT_UNLIKELY (meltxargdescr_ == MELTPAR_MARKGGC))
    {				/*mark for ggc */
      int meltix;
      meltframptr_ =
	(struct meltframe_meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE_st *)
	meltfirstargp_;
      /* use arguments meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE output_curframe_declstruct_init */
      (void) meltclosp_;
      (void) meltfirstargp_;
      (void) meltxargdescr_;
      (void) meltxargtab_;
      (void) meltxresdescr_;
      (void) meltxrestab_;
      /* marking closure */
      gt_ggc_mx_melt_un (meltframptr_->mcfr_clos);
      for (meltix = 0; meltix < 49; meltix++)
	if (meltframptr_->mcfr_varptr[meltix])
	  gt_ggc_mx_melt_un (meltframptr_->mcfr_varptr[meltix]);
      return NULL;
    } /*end markggc */ ;
  memset (&meltfram__, 0, sizeof (meltfram__));
/* declstructinit plain routine meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE nbval 49*/
  meltfram__.mcfr_nbvar = 49 /*nbval */ ;
  meltfram__.mcfr_clos = meltclosp_;
  meltfram__.mcfr_prev = (struct melt_callframe_st *) melt_topframe;
  melt_topframe = (struct melt_callframe_st *) &meltfram__;
  melt_trace_start ("NORMEXP_QUOTE", meltcallcount);
/*getargs*/

  /*getarg#0 */
  MELT_LOCATION ("warmelt-normal.melt:6447:/ getarg");
 /*_.RECV__V2*/ meltfptr[1] = (melt_ptr_t) meltfirstargp_;

  /*getarg#1 */
  /*^getarg */
  if (meltxargdescr_[0] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.ENV__V3*/ meltfptr[2] =
    (meltxargtab_[0].meltbp_aptr) ? (*(meltxargtab_[0].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2])) != NULL);


  /*getarg#2 */
  /*^getarg */
  if (meltxargdescr_[1] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.NCX__V4*/ meltfptr[3] =
    (meltxargtab_[1].meltbp_aptr) ? (*(meltxargtab_[1].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3])) != NULL);


  /*getarg#3 */
  /*^getarg */
  if (meltxargdescr_[2] != MELTBPAR_PTR)
    goto meltlab_endgetargs;
 /*_.PSLOC__V5*/ meltfptr[4] =
    (meltxargtab_[2].meltbp_aptr) ? (*(meltxargtab_[2].meltbp_aptr)) : NULL;
  gcc_assert (melt_discr ((melt_ptr_t) ( /*_.PSLOC__V5*/ meltfptr[4])) !=
	      NULL);

  ;
  goto meltlab_endgetargs;
meltlab_endgetargs:;
/*body*/
/*^block*/
/*anyblock*/
  {


#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6448:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L1*/ meltfnum[0] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6448:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L1*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L2*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6448:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[5];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L2*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6448;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_quote recv=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.RECV__V2*/ meltfptr[1];
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V6*/ meltfptr[5] =
	      /*_.MELT_DEBUG_FUN__V7*/ meltfptr[6];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6448:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L2*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V7*/ meltfptr[6] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V6*/ meltfptr[5] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6448:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L1*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IF___V6*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6449:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L3*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]),
			     (melt_ptr_t) (( /*!CLASS_SOURCE_QUOTE */
					    meltfrout->tabval[1])));;
      MELT_LOCATION ("warmelt-normal.melt:6449:/ cond");
      /*cond */ if ( /*_#IS_A__L3*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V9*/ meltfptr[5] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6449:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check quote recv"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6449) ? (6449) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V8*/ meltfptr[6] = /*_._IFELSE___V9*/ meltfptr[5];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6449:/ clear");
	     /*clear *//*_#IS_A__L3*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V9*/ meltfptr[5] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V8*/ meltfptr[6] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6450:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L4*/ meltfnum[0] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.ENV__V3*/ meltfptr[2]),
			     (melt_ptr_t) (( /*!CLASS_ENVIRONMENT */
					    meltfrout->tabval[2])));;
      MELT_LOCATION ("warmelt-normal.melt:6450:/ cond");
      /*cond */ if ( /*_#IS_A__L4*/ meltfnum[0])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V11*/ meltfptr[10] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6450:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check env"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6450) ? (6450) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V10*/ meltfptr[5] = /*_._IFELSE___V11*/ meltfptr[10];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6450:/ clear");
	     /*clear *//*_#IS_A__L4*/ meltfnum[0] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V11*/ meltfptr[10] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V10*/ meltfptr[5] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6451:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      MELT_CHECK_SIGNAL ();
      ;
   /*_#IS_A__L5*/ meltfnum[1] =
	melt_is_instance_of ((melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]),
			     (melt_ptr_t) (( /*!CLASS_NORMALIZATION_CONTEXT */
					    meltfrout->tabval[3])));;
      MELT_LOCATION ("warmelt-normal.melt:6451:/ cond");
      /*cond */ if ( /*_#IS_A__L5*/ meltfnum[1])	/*then */
	{
	  /*^cond.then */
	  /*_._IFELSE___V13*/ meltfptr[12] = ( /*nil */ NULL);;
	}
      else
	{
	  MELT_LOCATION ("warmelt-normal.melt:6451:/ cond.else");

	  /*^block */
	  /*anyblock */
	  {




	    {
	      /*^locexp */
	      melt_assert_failed (("check nctxt"),
				  ("warmelt-normal.melt")
				  ? ("warmelt-normal.melt") : __FILE__,
				  (6451) ? (6451) : __LINE__, __FUNCTION__);
	      ;
	    }
	    ;
	       /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
	    /*epilog */
	  }
	  ;
	}
      ;
      /*^compute */
      /*_.IFCPP___V12*/ meltfptr[10] = /*_._IFELSE___V13*/ meltfptr[12];;
      /*epilog */

      MELT_LOCATION ("warmelt-normal.melt:6451:/ clear");
	     /*clear *//*_#IS_A__L5*/ meltfnum[1] = 0;
      /*^clear */
	     /*clear *//*_._IFELSE___V13*/ meltfptr[12] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*_.IFCPP___V12*/ meltfptr[10] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
    ;

    {
      MELT_LOCATION ("warmelt-normal.melt:6452:/ locexp");

#if MELT_HAVE_DEBUG
      if (melt_need_debug (0))
	melt_dbgshortbacktrace (("normexp_quote"), (10));
#endif
      ;
    }
    ;
    MELT_LOCATION ("warmelt-normal.melt:6453:/ quasiblock");


    /*^getslot */
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 1, "LOCA_LOCATION");
  /*_.SLOC__V15*/ meltfptr[14] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6454:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.RECV__V2*/ meltfptr[1]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 2, "SQUOTED");
  /*_.QUOTED__V16*/ meltfptr[15] = slot;
    };
    ;
    MELT_LOCATION ("warmelt-normal.melt:6455:/ getslot");
    {
      melt_ptr_t slot = NULL, obj = NULL;
      obj = (melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]) /*=obj*/ ;
      melt_object_get_field (slot, obj, 10, "NCTX_CURPROC");
  /*_.CURPROC__V17*/ meltfptr[16] = slot;
    };
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_.DISCRIM__V18*/ meltfptr[17] =
      ((melt_ptr_t)
       (melt_discr ((melt_ptr_t) ( /*_.QUOTED__V16*/ meltfptr[15]))));;
    /*^compute */
 /*_#eqeq__L6*/ meltfnum[0] =
      (( /*_.DISCRIM__V18*/ meltfptr[17]) ==
       (( /*!CLASS_SYMBOL */ meltfrout->tabval[4])));;
    MELT_LOCATION ("warmelt-normal.melt:6458:/ cond");
    /*cond */ if ( /*_#eqeq__L6*/ meltfnum[0])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{


	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6459:/ apply");
	  /*apply */
	  {
	    union meltparam_un argtab[2];
	    memset (&argtab, 0, sizeof (argtab));
	    /*^apply.arg */
	    argtab[0].meltbp_aptr =
	      (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
	    /*^apply.arg */
	    argtab[1].meltbp_aptr =
	      (melt_ptr_t *) & /*_.PSLOC__V5*/ meltfptr[4];
	    /*_.NORMAL_SYMBOL_DATA__V20*/ meltfptr[19] =
	      melt_apply ((meltclosure_ptr_t)
			  (( /*!NORMAL_SYMBOL_DATA */ meltfrout->tabval[5])),
			  (melt_ptr_t) ( /*_.QUOTED__V16*/ meltfptr[15]),
			  (MELTBPARSTR_PTR MELTBPARSTR_PTR ""), argtab, "",
			  (union meltparam_un *) 0);
	  }
	  ;
	  /*_.NDATA__V19*/ meltfptr[18] =
	    /*_.NORMAL_SYMBOL_DATA__V20*/ meltfptr[19];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6458:/ clear");
	     /*clear *//*_.NORMAL_SYMBOL_DATA__V20*/ meltfptr[19] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

	/*^block */
	/*anyblock */
	{


	  MELT_CHECK_SIGNAL ();
	  ;
   /*_#IS_INTEGERBOX__L7*/ meltfnum[1] =
	    (melt_magic_discr ((melt_ptr_t) ( /*_.QUOTED__V16*/ meltfptr[15]))
	     == MELTOBMAG_INT);;
	  MELT_LOCATION ("warmelt-normal.melt:6460:/ cond");
	  /*cond */ if ( /*_#IS_INTEGERBOX__L7*/ meltfnum[1])	/*then */
	    {
	      /*^cond.then */
	      /*^block */
	      /*anyblock */
	      {

		MELT_LOCATION ("warmelt-normal.melt:6461:/ quasiblock");



		MELT_CHECK_SIGNAL ();
		;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6463:/ apply");
		/*apply */
		{
		  union meltparam_un argtab[3];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^apply.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
		  /*^apply.arg */
		  argtab[1].meltbp_aptr =
		    (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[14];
		  /*^apply.arg */
		  argtab[2].meltbp_cstring = "discr_constant_integer";
		  /*_.NORMAL_PREDEF__V23*/ meltfptr[22] =
		    melt_apply ((meltclosure_ptr_t)
				(( /*!NORMAL_PREDEF */ meltfrout->tabval[7])),
				(melt_ptr_t) (( /*!DISCR_CONSTANT_INTEGER */
					       meltfrout->tabval[8])),
				(MELTBPARSTR_PTR MELTBPARSTR_PTR
				 MELTBPARSTR_CSTRING ""), argtab, "",
				(union meltparam_un *) 0);
		}
		;
		MELT_LOCATION ("warmelt-normal.melt:6462:/ quasiblock");


		/*^rawallocobj */
		/*rawallocobj */
		{
		  melt_ptr_t newobj = 0;
		  melt_raw_object_create (newobj,
					  (melt_ptr_t) (( /*!CLASS_NREP_DATABOXEDINTEGER */ meltfrout->tabval[6])), (6), "CLASS_NREP_DATABOXEDINTEGER");
      /*_.INST__V25*/ meltfptr[24] =
		    newobj;
		};
		;
		/*^putslot */
		/*putslot */
		melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
				melt_magic_discr ((melt_ptr_t)
						  ( /*_.INST__V25*/
						   meltfptr[24])) ==
				MELTOBMAG_OBJECT);
		melt_putfield_object (( /*_.INST__V25*/ meltfptr[24]), (4),
				      ( /*_.NORMAL_PREDEF__V23*/
				       meltfptr[22]), "NDATA_DISCRX");
		;
		/*^putslot */
		/*putslot */
		melt_assertmsg ("checkobj putslot  _ @NBOXINT_NUM",
				melt_magic_discr ((melt_ptr_t)
						  ( /*_.INST__V25*/
						   meltfptr[24])) ==
				MELTOBMAG_OBJECT);
		melt_putfield_object (( /*_.INST__V25*/ meltfptr[24]), (5),
				      ( /*_.QUOTED__V16*/ meltfptr[15]),
				      "NBOXINT_NUM");
		;
		/*^touchobj */

		melt_dbgtrace_written_object ( /*_.INST__V25*/ meltfptr[24],
					      "newly made instance");
		;
		/*_.NINTDATA__V24*/ meltfptr[23] =
		  /*_.INST__V25*/ meltfptr[24];;

		MELT_CHECK_SIGNAL ();
		;
		MELT_LOCATION ("warmelt-normal.melt:6466:/ apply");
		/*apply */
		{
		  union meltparam_un argtab[1];
		  memset (&argtab, 0, sizeof (argtab));
		  /*^apply.arg */
		  argtab[0].meltbp_aptr =
		    (melt_ptr_t *) & /*_.NINTDATA__V24*/ meltfptr[23];
		  /*_.ADD_NCTX_DATA__V26*/ meltfptr[25] =
		    melt_apply ((meltclosure_ptr_t)
				(( /*!ADD_NCTX_DATA */ meltfrout->tabval[9])),
				(melt_ptr_t) ( /*_.NCX__V4*/ meltfptr[3]),
				(MELTBPARSTR_PTR ""), argtab, "",
				(union meltparam_un *) 0);
		}
		;
		/*_.LET___V22*/ meltfptr[21] =
		  /*_.NINTDATA__V24*/ meltfptr[23];;

		MELT_LOCATION ("warmelt-normal.melt:6461:/ clear");
	       /*clear *//*_.NORMAL_PREDEF__V23*/ meltfptr[22] = 0;
		/*^clear */
	       /*clear *//*_.NINTDATA__V24*/ meltfptr[23] = 0;
		/*^clear */
	       /*clear *//*_.ADD_NCTX_DATA__V26*/ meltfptr[25] = 0;
		/*_._IFELSE___V21*/ meltfptr[19] =
		  /*_.LET___V22*/ meltfptr[21];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6460:/ clear");
	       /*clear *//*_.LET___V22*/ meltfptr[21] = 0;
	      }
	      ;
	    }
	  else
	    {			/*^cond.else */

	      /*^block */
	      /*anyblock */
	      {


		MELT_CHECK_SIGNAL ();
		;
     /*_#IS_STRING__L8*/ meltfnum[7] =
		  (melt_magic_discr
		   ((melt_ptr_t) ( /*_.QUOTED__V16*/ meltfptr[15])) ==
		   MELTOBMAG_STRING);;
		MELT_LOCATION ("warmelt-normal.melt:6469:/ cond");
		/*cond */ if ( /*_#IS_STRING__L8*/ meltfnum[7])	/*then */
		  {
		    /*^cond.then */
		    /*^block */
		    /*anyblock */
		    {

		      MELT_LOCATION ("warmelt-normal.melt:6470:/ quasiblock");



		      MELT_CHECK_SIGNAL ();
		      ;

		      MELT_CHECK_SIGNAL ();
		      ;
		      MELT_LOCATION ("warmelt-normal.melt:6472:/ apply");
		      /*apply */
		      {
			union meltparam_un argtab[3];
			memset (&argtab, 0, sizeof (argtab));
			/*^apply.arg */
			argtab[0].meltbp_aptr =
			  (melt_ptr_t *) & /*_.NCX__V4*/ meltfptr[3];
			/*^apply.arg */
			argtab[1].meltbp_aptr =
			  (melt_ptr_t *) & /*_.SLOC__V15*/ meltfptr[14];
			/*^apply.arg */
			argtab[2].meltbp_cstring = "discr_string";
			/*_.NORMAL_PREDEF__V29*/ meltfptr[25] =
			  melt_apply ((meltclosure_ptr_t)
				      (( /*!NORMAL_PREDEF */ meltfrout->
					tabval[7])),
				      (melt_ptr_t) (( /*!DISCR_STRING */
						     meltfrout->tabval[11])),
				      (MELTBPARSTR_PTR MELTBPARSTR_PTR
				       MELTBPARSTR_CSTRING ""), argtab, "",
				      (union meltparam_un *) 0);
		      }
		      ;
		      MELT_LOCATION ("warmelt-normal.melt:6471:/ quasiblock");


		      /*^rawallocobj */
		      /*rawallocobj */
		      {
			melt_ptr_t newobj = 0;
			melt_raw_object_create (newobj,
						(melt_ptr_t) (( /*!CLASS_NREP_DATASTRING */ meltfrout->tabval[10])), (6), "CLASS_NREP_DATASTRING");
	/*_.INST__V31*/ meltfptr[30] =
			  newobj;
		      };
		      ;
		      /*^putslot */
		      /*putslot */
		      melt_assertmsg ("checkobj putslot  _ @NDATA_DISCRX",
				      melt_magic_discr ((melt_ptr_t)
							( /*_.INST__V31*/
							 meltfptr[30])) ==
				      MELTOBMAG_OBJECT);
		      melt_putfield_object (( /*_.INST__V31*/ meltfptr[30]),
					    (4),
					    ( /*_.NORMAL_PREDEF__V29*/
					     meltfptr[25]), "NDATA_DISCRX");
		      ;
		      /*^putslot */
		      /*putslot */
		      melt_assertmsg ("checkobj putslot  _ @NSTR_STRING",
				      melt_magic_discr ((melt_ptr_t)
							( /*_.INST__V31*/
							 meltfptr[30])) ==
				      MELTOBMAG_OBJECT);
		      melt_putfield_object (( /*_.INST__V31*/ meltfptr[30]),
					    (5),
					    ( /*_.QUOTED__V16*/ meltfptr[15]),
					    "NSTR_STRING");
		      ;
		      /*^touchobj */

		      melt_dbgtrace_written_object ( /*_.INST__V31*/
						    meltfptr[30],
						    "newly made instance");
		      ;
		      /*_.NSTRDATA__V30*/ meltfptr[21] =
			/*_.INST__V31*/ meltfptr[30];;

		      MELT_CHECK_SIGNAL ();
		      ;
		      MELT_LOCATION ("warmelt-normal.melt:6475:/ apply");
		      /*apply */
		      {
			union meltparam_un argtab[1];
			memset (&argtab, 0, sizeof (argtab));
			/*^apply.arg */
			argtab[0].meltbp_aptr =
			  (melt_ptr_t *) & /*_.NSTRDATA__V30*/ meltfptr[21];
			/*_.ADD_NCTX_DATA__V32*/ meltfptr[31] =
			  melt_apply ((meltclosure_ptr_t)
				      (( /*!ADD_NCTX_DATA */ meltfrout->
					tabval[9])),
				      (melt_ptr_t) ( /*_.NCX__V4*/
						    meltfptr[3]),
				      (MELTBPARSTR_PTR ""), argtab, "",
				      (union meltparam_un *) 0);
		      }
		      ;
		      /*_.LET___V28*/ meltfptr[23] =
			/*_.NSTRDATA__V30*/ meltfptr[21];;

		      MELT_LOCATION ("warmelt-normal.melt:6470:/ clear");
		 /*clear *//*_.NORMAL_PREDEF__V29*/ meltfptr[25] = 0;
		      /*^clear */
		 /*clear *//*_.NSTRDATA__V30*/ meltfptr[21] = 0;
		      /*^clear */
		 /*clear *//*_.ADD_NCTX_DATA__V32*/ meltfptr[31] = 0;
		      /*_._IFELSE___V27*/ meltfptr[22] =
			/*_.LET___V28*/ meltfptr[23];;
		      /*epilog */

		      MELT_LOCATION ("warmelt-normal.melt:6469:/ clear");
		 /*clear *//*_.LET___V28*/ meltfptr[23] = 0;
		    }
		    ;
		  }
		else
		  {		/*^cond.else */

		    /*^block */
		    /*anyblock */
		    {


#if MELT_HAVE_DEBUG
		      MELT_LOCATION ("warmelt-normal.melt:6480:/ cppif.then");
		      /*^block */
		      /*anyblock */
		      {


			MELT_CHECK_SIGNAL ();
			;
			/*^cond */
			/*cond */ if (( /*nil */ NULL))	/*then */
			  {
			    /*^cond.then */
			    /*_._IFELSE___V34*/ meltfptr[21] =
			      ( /*nil */ NULL);;
			  }
			else
			  {
			    MELT_LOCATION
			      ("warmelt-normal.melt:6480:/ cond.else");

			    /*^block */
			    /*anyblock */
			    {




			      {
				/*^locexp */
				melt_assert_failed (("unexpected quoted stuff"), ("warmelt-normal.melt") ? ("warmelt-normal.melt") : __FILE__, (6480) ? (6480) : __LINE__, __FUNCTION__);
				;
			      }
			      ;
		     /*clear *//*_._IFELSE___V34*/ meltfptr[21] = 0;
			      /*epilog */
			    }
			    ;
			  }
			;
			/*^compute */
			/*_.IFCPP___V33*/ meltfptr[25] =
			  /*_._IFELSE___V34*/ meltfptr[21];;
			/*epilog */

			MELT_LOCATION ("warmelt-normal.melt:6480:/ clear");
		   /*clear *//*_._IFELSE___V34*/ meltfptr[21] = 0;
		      }

#else /*MELT_HAVE_DEBUG */
		      /*^cppif.else */
		      /*_.IFCPP___V33*/ meltfptr[25] = ( /*nil */ NULL);
#endif /*MELT_HAVE_DEBUG */
		      ;
		      MELT_LOCATION ("warmelt-normal.melt:6477:/ quasiblock");


		      /*_.PROGN___V35*/ meltfptr[31] =
			/*_.IFCPP___V33*/ meltfptr[25];;
		      /*^compute */
		      /*_._IFELSE___V27*/ meltfptr[22] =
			/*_.PROGN___V35*/ meltfptr[31];;
		      /*epilog */

		      MELT_LOCATION ("warmelt-normal.melt:6469:/ clear");
		 /*clear *//*_.IFCPP___V33*/ meltfptr[25] = 0;
		      /*^clear */
		 /*clear *//*_.PROGN___V35*/ meltfptr[31] = 0;
		    }
		    ;
		  }
		;
		/*_._IFELSE___V21*/ meltfptr[19] =
		  /*_._IFELSE___V27*/ meltfptr[22];;
		/*epilog */

		MELT_LOCATION ("warmelt-normal.melt:6460:/ clear");
	       /*clear *//*_#IS_STRING__L8*/ meltfnum[7] = 0;
		/*^clear */
	       /*clear *//*_._IFELSE___V27*/ meltfptr[22] = 0;
	      }
	      ;
	    }
	  ;
	  /*_.NDATA__V19*/ meltfptr[18] = /*_._IFELSE___V21*/ meltfptr[19];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6458:/ clear");
	     /*clear *//*_#IS_INTEGERBOX__L7*/ meltfnum[1] = 0;
	  /*^clear */
	     /*clear *//*_._IFELSE___V21*/ meltfptr[19] = 0;
	}
	;
      }
    ;

#if MELT_HAVE_DEBUG
    MELT_LOCATION ("warmelt-normal.melt:6484:/ cppif.then");
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	melt_dbgcounter++;
#endif
	;
      }
      ;

      MELT_CHECK_SIGNAL ();
      ;
   /*_#MELT_NEED_DBG__L9*/ meltfnum[7] =
	/*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	;;
      MELT_LOCATION ("warmelt-normal.melt:6484:/ cond");
      /*cond */ if ( /*_#MELT_NEED_DBG__L9*/ meltfnum[7])	/*then */
	{
	  /*^cond.then */
	  /*^block */
	  /*anyblock */
	  {

     /*_#MELT_CALLCOUNT__L10*/ meltfnum[1] =
	      /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
	      meltcallcount	/* melt_callcount debugging */
#else
	      0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
	      ;;

	    MELT_CHECK_SIGNAL ();
	    ;
	    MELT_LOCATION ("warmelt-normal.melt:6484:/ apply");
	    /*apply */
	    {
	      union meltparam_un argtab[7];
	      memset (&argtab, 0, sizeof (argtab));
	      /*^apply.arg */
	      argtab[0].meltbp_long = /*_#MELT_CALLCOUNT__L10*/ meltfnum[1];
	      /*^apply.arg */
	      argtab[1].meltbp_cstring = "warmelt-normal.melt";
	      /*^apply.arg */
	      argtab[2].meltbp_long = 6484;
	      /*^apply.arg */
	      argtab[3].meltbp_cstring = "normexp_quote ndata=";
	      /*^apply.arg */
	      argtab[4].meltbp_aptr =
		(melt_ptr_t *) & /*_.NDATA__V19*/ meltfptr[18];
	      /*^apply.arg */
	      argtab[5].meltbp_cstring = "\n.. curproc=";
	      /*^apply.arg */
	      argtab[6].meltbp_aptr =
		(melt_ptr_t *) & /*_.CURPROC__V17*/ meltfptr[16];
	      /*_.MELT_DEBUG_FUN__V37*/ meltfptr[21] =
		melt_apply ((meltclosure_ptr_t)
			    (( /*!MELT_DEBUG_FUN */ meltfrout->tabval[0])),
			    (melt_ptr_t) (( /*nil */ NULL)),
			    (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_LONG MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR MELTBPARSTR_CSTRING
			     MELTBPARSTR_PTR ""), argtab, "",
			    (union meltparam_un *) 0);
	    }
	    ;
	    /*_._IF___V36*/ meltfptr[23] =
	      /*_.MELT_DEBUG_FUN__V37*/ meltfptr[21];;
	    /*epilog */

	    MELT_LOCATION ("warmelt-normal.melt:6484:/ clear");
	       /*clear *//*_#MELT_CALLCOUNT__L10*/ meltfnum[1] = 0;
	    /*^clear */
	       /*clear *//*_.MELT_DEBUG_FUN__V37*/ meltfptr[21] = 0;
	  }
	  ;
	}
      else
	{			/*^cond.else */

    /*_._IF___V36*/ meltfptr[23] = NULL;;
	}
      ;

      {
	MELT_LOCATION ("warmelt-normal.melt:6484:/ locexp");
	/*void */ (void) 0;
      }
      ;
      /*^quasiblock */


      /*epilog */

      /*^clear */
	     /*clear *//*_#MELT_NEED_DBG__L9*/ meltfnum[7] = 0;
      /*^clear */
	     /*clear *//*_._IF___V36*/ meltfptr[23] = 0;
    }

#else /*MELT_HAVE_DEBUG */
    /*^cppif.else */
    /*^block */
    /*anyblock */
    {


      {
	/*^locexp */
	/*void */ (void) 0;
      }
      ;
      /*epilog */
    }

#endif /*MELT_HAVE_DEBUG */
    ;

    MELT_CHECK_SIGNAL ();
    ;
 /*_#IS_A__L11*/ meltfnum[1] =
      melt_is_instance_of ((melt_ptr_t) ( /*_.CURPROC__V17*/ meltfptr[16]),
			   (melt_ptr_t) (( /*!CLASS_NREP_ROUTPROC */
					  meltfrout->tabval[12])));;
    MELT_LOCATION ("warmelt-normal.melt:6485:/ cond");
    /*cond */ if ( /*_#IS_A__L11*/ meltfnum[1])	/*then */
      {
	/*^cond.then */
	/*^block */
	/*anyblock */
	{

	  MELT_LOCATION ("warmelt-normal.melt:6486:/ quasiblock");



	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6487:/ quasiblock");


	  /*^rawallocobj */
	  /*rawallocobj */
	  {
	    melt_ptr_t newobj = 0;
	    melt_raw_object_create (newobj,
				    (melt_ptr_t) (( /*!CLASS_NREP_CONSTANT */
						   meltfrout->tabval[13])),
				    (4), "CLASS_NREP_CONSTANT");
    /*_.INST__V41*/ meltfptr[19] =
	      newobj;
	  };
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NREP_LOC",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V41*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V41*/ meltfptr[19]), (0),
				( /*_.SLOC__V15*/ meltfptr[14]), "NREP_LOC");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NCONST_SVAL",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V41*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V41*/ meltfptr[19]), (1),
				( /*_.QUOTED__V16*/ meltfptr[15]),
				"NCONST_SVAL");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NCONST_DATA",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V41*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V41*/ meltfptr[19]), (3),
				( /*_.NDATA__V19*/ meltfptr[18]),
				"NCONST_DATA");
	  ;
	  /*^putslot */
	  /*putslot */
	  melt_assertmsg ("checkobj putslot  _ @NCONST_PROC",
			  melt_magic_discr ((melt_ptr_t)
					    ( /*_.INST__V41*/ meltfptr[19]))
			  == MELTOBMAG_OBJECT);
	  melt_putfield_object (( /*_.INST__V41*/ meltfptr[19]), (2),
				( /*_.CURPROC__V17*/ meltfptr[16]),
				"NCONST_PROC");
	  ;
	  /*^touchobj */

	  melt_dbgtrace_written_object ( /*_.INST__V41*/ meltfptr[19],
					"newly made instance");
	  ;
	  /*_.NCONST__V40*/ meltfptr[22] = /*_.INST__V41*/ meltfptr[19];;
	  MELT_LOCATION ("warmelt-normal.melt:6493:/ getslot");
	  {
	    melt_ptr_t slot = NULL, obj = NULL;
	    obj = (melt_ptr_t) ( /*_.CURPROC__V17*/ meltfptr[16]) /*=obj*/ ;
	    melt_object_get_field (slot, obj, 5, "NRPRO_CONST");
    /*_.NRPRO_CONST__V42*/ meltfptr[21] = slot;
	  };
	  ;

	  {
	    /*^locexp */
	    meltgc_append_list ((melt_ptr_t)
				( /*_.NRPRO_CONST__V42*/ meltfptr[21]),
				(melt_ptr_t) ( /*_.NDATA__V19*/
					      meltfptr[18]));
	  }
	  ;

#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6494:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L12*/ meltfnum[7] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6494:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L12*/ meltfnum[7])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L13*/ meltfnum[12] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6494:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L13*/ meltfnum[12];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6494;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring =
		      "normexp_quote in routine nconst=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NCONST__V40*/ meltfptr[22];
		    /*_.MELT_DEBUG_FUN__V44*/ meltfptr[43] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V43*/ meltfptr[23] =
		    /*_.MELT_DEBUG_FUN__V44*/ meltfptr[43];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6494:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L13*/ meltfnum[12] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V44*/ meltfptr[43] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V43*/ meltfptr[23] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6494:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L12*/ meltfnum[7] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V43*/ meltfptr[23] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6495:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.NCONST__V40*/ meltfptr[22];;
	  MELT_LOCATION ("warmelt-normal.melt:6495:/ putxtraresult");
	  if (!meltxrestab_ || !meltxresdescr_)
	    goto meltlabend_rout;
	  if (meltxresdescr_[0] != MELTBPAR_PTR)
	    goto meltlabend_rout;
	  if (meltxrestab_[0].meltbp_aptr)
	    *(meltxrestab_[0].meltbp_aptr) = (melt_ptr_t) (( /*nil */ NULL));
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  /*_.LET___V39*/ meltfptr[31] = /*_.RETURN___V45*/ meltfptr[43];;

	  MELT_LOCATION ("warmelt-normal.melt:6486:/ clear");
	     /*clear *//*_.NCONST__V40*/ meltfptr[22] = 0;
	  /*^clear */
	     /*clear *//*_.NRPRO_CONST__V42*/ meltfptr[21] = 0;
	  /*^clear */
	     /*clear *//*_.RETURN___V45*/ meltfptr[43] = 0;
	  /*_._IFELSE___V38*/ meltfptr[25] = /*_.LET___V39*/ meltfptr[31];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6485:/ clear");
	     /*clear *//*_.LET___V39*/ meltfptr[31] = 0;
	}
	;
      }
    else
      {				/*^cond.else */

	/*^block */
	/*anyblock */
	{


#if MELT_HAVE_DEBUG
	  MELT_LOCATION ("warmelt-normal.melt:6497:/ cppif.then");
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*melt_increment_dbgcounter */
#if MELT_HAVE_DEBUG
	      melt_dbgcounter++;
#endif
	      ;
	    }
	    ;

	    MELT_CHECK_SIGNAL ();
	    ;
     /*_#MELT_NEED_DBG__L14*/ meltfnum[12] =
	      /*MELT_NEED_DBG */
#if MELT_HAVE_DEBUG
	      ( /*melt_need_dbg */ melt_need_debug ((int) 0))
#else
	      0			/* no melt_need_dbg */
#endif /*MELT_HAVE_DEBUG */
	      ;;
	    MELT_LOCATION ("warmelt-normal.melt:6497:/ cond");
	    /*cond */ if ( /*_#MELT_NEED_DBG__L14*/ meltfnum[12])	/*then */
	      {
		/*^cond.then */
		/*^block */
		/*anyblock */
		{

       /*_#MELT_CALLCOUNT__L15*/ meltfnum[7] =
		    /* for melt_callcount */
#if MELT_HAVE_DEBUG && defined (meltcallcount)
		    meltcallcount	/* melt_callcount debugging */
#else
		    0L		/* melt_callcount without debug */
#endif /* MELT_HAVE_DEBUG melt_callcount */
		    ;;

		  MELT_CHECK_SIGNAL ();
		  ;
		  MELT_LOCATION ("warmelt-normal.melt:6497:/ apply");
		  /*apply */
		  {
		    union meltparam_un argtab[5];
		    memset (&argtab, 0, sizeof (argtab));
		    /*^apply.arg */
		    argtab[0].meltbp_long =
		      /*_#MELT_CALLCOUNT__L15*/ meltfnum[7];
		    /*^apply.arg */
		    argtab[1].meltbp_cstring = "warmelt-normal.melt";
		    /*^apply.arg */
		    argtab[2].meltbp_long = 6497;
		    /*^apply.arg */
		    argtab[3].meltbp_cstring = "normexp_quote in init ndata=";
		    /*^apply.arg */
		    argtab[4].meltbp_aptr =
		      (melt_ptr_t *) & /*_.NDATA__V19*/ meltfptr[18];
		    /*_.MELT_DEBUG_FUN__V47*/ meltfptr[22] =
		      melt_apply ((meltclosure_ptr_t)
				  (( /*!MELT_DEBUG_FUN */ meltfrout->
				    tabval[0])),
				  (melt_ptr_t) (( /*nil */ NULL)),
				  (MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_LONG MELTBPARSTR_CSTRING
				   MELTBPARSTR_PTR ""), argtab, "",
				  (union meltparam_un *) 0);
		  }
		  ;
		  /*_._IF___V46*/ meltfptr[23] =
		    /*_.MELT_DEBUG_FUN__V47*/ meltfptr[22];;
		  /*epilog */

		  MELT_LOCATION ("warmelt-normal.melt:6497:/ clear");
		 /*clear *//*_#MELT_CALLCOUNT__L15*/ meltfnum[7] = 0;
		  /*^clear */
		 /*clear *//*_.MELT_DEBUG_FUN__V47*/ meltfptr[22] = 0;
		}
		;
	      }
	    else
	      {			/*^cond.else */

      /*_._IF___V46*/ meltfptr[23] = NULL;;
	      }
	    ;

	    {
	      MELT_LOCATION ("warmelt-normal.melt:6497:/ locexp");
	      /*void */ (void) 0;
	    }
	    ;
	    /*^quasiblock */


	    /*epilog */

	    /*^clear */
	       /*clear *//*_#MELT_NEED_DBG__L14*/ meltfnum[12] = 0;
	    /*^clear */
	       /*clear *//*_._IF___V46*/ meltfptr[23] = 0;
	  }

#else /*MELT_HAVE_DEBUG */
	  /*^cppif.else */
	  /*^block */
	  /*anyblock */
	  {


	    {
	      /*^locexp */
	      /*void */ (void) 0;
	    }
	    ;
	    /*epilog */
	  }

#endif /*MELT_HAVE_DEBUG */
	  ;

	  MELT_CHECK_SIGNAL ();
	  ;
	  MELT_LOCATION ("warmelt-normal.melt:6498:/ quasiblock");


	  /*_._RETVAL___V1*/ meltfptr[0] = /*_.NDATA__V19*/ meltfptr[18];;
	  MELT_LOCATION ("warmelt-normal.melt:6498:/ putxtraresult");
	  if (!meltxrestab_ || !meltxresdescr_)
	    goto meltlabend_rout;
	  if (meltxresdescr_[0] != MELTBPAR_PTR)
	    goto meltlabend_rout;
	  if (meltxrestab_[0].meltbp_aptr)
	    *(meltxrestab_[0].meltbp_aptr) = (melt_ptr_t) (( /*nil */ NULL));
	  ;
	  /*^finalreturn */
	  ;
	  /*finalret */ goto meltlabend_rout;
	  MELT_LOCATION ("warmelt-normal.melt:6496:/ quasiblock");


	  /*_.PROGN___V49*/ meltfptr[43] = /*_.RETURN___V48*/ meltfptr[21];;
	  /*^compute */
	  /*_._IFELSE___V38*/ meltfptr[25] = /*_.PROGN___V49*/ meltfptr[43];;
	  /*epilog */

	  MELT_LOCATION ("warmelt-normal.melt:6485:/ clear");
	     /*clear *//*_.RETURN___V48*/ meltfptr[21] = 0;
	  /*^clear */
	     /*clear *//*_.PROGN___V49*/ meltfptr[43] = 0;
	}
	;
      }
    ;
    /*_.LET___V14*/ meltfptr[12] = /*_._IFELSE___V38*/ meltfptr[25];;

    MELT_LOCATION ("warmelt-normal.melt:6453:/ clear");
	   /*clear *//*_.SLOC__V15*/ meltfptr[14] = 0;
    /*^clear */
	   /*clear *//*_.QUOTED__V16*/ meltfptr[15] = 0;
    /*^clear */
	   /*clear *//*_.CURPROC__V17*/ meltfptr[16] = 0;
    /*^clear */
	   /*clear *//*_.DISCRIM__V18*/ meltfptr[17] = 0;
    /*^clear */
	   /*clear *//*_#eqeq__L6*/ meltfnum[0] = 0;
    /*^clear */
	   /*clear *//*_.NDATA__V19*/ meltfptr[18] = 0;
    /*^clear */
	   /*clear *//*_#IS_A__L11*/ meltfnum[1] = 0;
    /*^clear */
	   /*clear *//*_._IFELSE___V38*/ meltfptr[25] = 0;

    MELT_CHECK_SIGNAL ();
    ;
    MELT_LOCATION ("warmelt-normal.melt:6447:/ quasiblock");


    /*_._RETVAL___V1*/ meltfptr[0] = /*_.LET___V14*/ meltfptr[12];;

    {
      MELT_LOCATION ("warmelt-normal.melt:6447:/ locexp");
      /*ochecknores compilobj_nrep_return */
#if MELT_HAVE_DEBUG
      if (meltxresdescr_ && meltxresdescr_[0] && meltxrestab_)
	melt_warn_for_no_expected_secondary_results ();
/* we warned when secondary results are expected but not returned. */
#endif /*MELT_HAVE_DEBUG */
      ;
    }
    ;
    /*^finalreturn */
    ;
    /*finalret */ goto meltlabend_rout;
    /*epilog */

    /*^clear */
	   /*clear *//*_.IFCPP___V8*/ meltfptr[6] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V10*/ meltfptr[5] = 0;
    /*^clear */
	   /*clear *//*_.IFCPP___V12*/ meltfptr[10] = 0;
    /*^clear */
	   /*clear *//*_.LET___V14*/ meltfptr[12] = 0;
  }

  ;
  goto meltlabend_rout;
meltlabend_rout:
  melt_trace_end ("NORMEXP_QUOTE", meltcallcount);
  MELT_TRACE_EXIT_LOCATION ();
  melt_blocklevel_signals =
    current_blocklevel_signals_meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE_melt;
  melt_topframe = (struct melt_callframe_st *) meltfram__.mcfr_prev;
  return (melt_ptr_t) ( /*_._RETVAL___V1*/ meltfptr[0]);
#undef meltcallcount
#undef meltfram__
#undef MELTFRAM_NBVARNUM
#undef MELTFRAM_NBVARPTR
}				/*end meltrout_120_WARMELTmiNORMAL_NORMEXP_QUOTE */



/**** end of warmelt-normal+04.c ****/
