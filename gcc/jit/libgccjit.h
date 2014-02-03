/*
  A pure C API to enable client code to embed GCC as a JIT-compiler.
 */
#ifndef LIBGCCJIT_H
#define LIBGCCJIT_H

#ifdef __cplusplus
extern "C" {
#endif /* __cplusplus */

/**********************************************************************
 Data structures.
 **********************************************************************/
/* All structs within the API are opaque. */

/* A gcc_jit_context encapsulates the state of a compilation.  It goes
   through two states:

   (1) "initial", during which you can set up options on it, and add
       types, functions and code, using the API below.
       Invoking gcc_jit_context_compile on it transitions it to the
       "after compilation" state.

   (2) "after compilation", when you can call gcc_jit_context_release to
       clean up.  */
typedef struct gcc_jit_context gcc_jit_context;

/* A gcc_jit_result encapsulates the result of a compilation.  */
typedef struct gcc_jit_result gcc_jit_result;

/* An object created within a context.  Such objects are automatically
   cleaned up when the context is released.

   The class hierarchy looks like this:

     +- gcc_jit_object
         +- gcc_jit_location
         +- gcc_jit_type
         +- gcc_jit_field
         +- gcc_jit_function
         +- gcc_jit_label
         +- gcc_jit_rvalue
             +- gcc_jit_lvalue
                 +- gcc_jit_param
         +- gcc_jit_loop
*/
typedef struct gcc_jit_object gcc_jit_object;

/* A gcc_jit_location encapsulates a source code location, so that
   you can (optionally) associate locations in your language with
   statements in the JIT-compiled code, allowing the debugger to
   single-step through your language.

   Note that to do so, you also need to enable
     GCC_JIT_BOOL_OPTION_DEBUGINFO
   on the gcc_jit_context.

   gcc_jit_location instances are optional; you can always pass
   NULL.  */
typedef struct gcc_jit_location gcc_jit_location;

/* A gcc_jit_type encapsulates a type e.g. "int" or a "struct foo*".  */
typedef struct gcc_jit_type gcc_jit_type;

/* A gcc_jit_field encapsulates a field within a struct; it is used
   when creating a struct type (using gcc_jit_context_new_struct_type).
   Fields cannot be shared between structs.  */
typedef struct gcc_jit_field gcc_jit_field;

/* A gcc_jit_function encapsulates a function: either one that you're
   creating yourself, or a reference to one that you're dynamically
   linking to within the rest of the process.  */
typedef struct gcc_jit_function gcc_jit_function;

/* A gcc_jit_label is a jump target within a function, for control flow.  */
typedef struct gcc_jit_label gcc_jit_label;

/* A gcc_jit_rvalue is an expression within your code, with some type.  */
typedef struct gcc_jit_rvalue gcc_jit_rvalue;

/* A gcc_jit_lvalue is a storage location within your code (e.g. a
   variable, a parameter, etc).  It is also a gcc_jit_rvalue; use
   gcc_jit_lvalue_as_rvalue to cast.  */
typedef struct gcc_jit_lvalue gcc_jit_lvalue;

/* A gcc_jit_param is a function parameter, used when creating a
   gcc_jit_function.  It is also a gcc_jit_lvalue (and thus also an
   rvalue); use gcc_jit_param_as_lvalue to convert.  */
typedef struct gcc_jit_param gcc_jit_param;

/* A gcc_jit_loop is a pre-canned way of creating loops without needing
   to manually manage gcc_jit_label instances.  */
typedef struct gcc_jit_loop gcc_jit_loop;

/*
   Acquire a JIT-compilation context.

   FIXME: error-handling?
*/
extern gcc_jit_context *
gcc_jit_context_acquire (void);

/* Release the context.  After this call, it's no longer valid to use
   the ctxt.  */
extern void
gcc_jit_context_release (gcc_jit_context *ctxt);

/* Options taking string values. */
enum gcc_jit_str_option
{
  /* The name of the program, for use as a prefix when printing error
     messages to stderr.  If NULL, or default, "libgccjit.so" is used.  */
  GCC_JIT_STR_OPTION_PROGNAME,

  GCC_JIT_NUM_STR_OPTIONS
};

/* Options taking int values. */
enum gcc_jit_int_option
{
  /* How much to optimize the code.
     Valid values are 0-3, corresponding to GCC's command-line options
     -O0 through -O3.

     The default value is 0 (unoptimized).  */
  GCC_JIT_INT_OPTION_OPTIMIZATION_LEVEL,

  GCC_JIT_NUM_INT_OPTIONS
};

/* Options taking boolean values.
   These all default to "false".  */
enum gcc_jit_bool_option
{
  /* If true, gcc_jit_context_compile will attempt to do the right
     thing so that if you attach a debugger to the process, it will
     be able to inspect variables and step through your code.

     Note that you can't step through code unless you set up source
     location information for the code (by creating and passing in
     gcc_jit_location instances).  */
  GCC_JIT_BOOL_OPTION_DEBUGINFO,

  /* If true, gcc_jit_context_compile will dump its initial "tree"
     representation of your code to stderr (before any
     optimizations).  */
  GCC_JIT_BOOL_OPTION_DUMP_INITIAL_TREE,

  /* If true, gcc_jit_context_compile will dump the "gimple"
     representation of your code to stderr, before any optimizations
     are performed.  The dump resembles C code.  */
  GCC_JIT_BOOL_OPTION_DUMP_INITIAL_GIMPLE,

  /* If true, gcc_jit_context_compile will dump the final
     generated code to stderr, in the form of assembly language.  */
  GCC_JIT_BOOL_OPTION_DUMP_GENERATED_CODE,

  /* If true, gcc_jit_context_compile will print information to stderr
     on the actions it is performing, followed by a profile showing
     the time taken and memory usage of each phase.
   */
  GCC_JIT_BOOL_OPTION_DUMP_SUMMARY,

  /* If true, gcc_jit_context_compile will dump copious
     amount of information on what it's doing to various
     files within a temporary directory.  Use
     GCC_JIT_BOOL_OPTION_KEEP_INTERMEDIATES (see below) to
     see the results.  The files are intended to be human-readable,
     but the exact files and their formats are subject to change.
  */
  GCC_JIT_BOOL_OPTION_DUMP_EVERYTHING,

  /* If true, libgccjit will aggressively run its garbage collector, to
     shake out bugs (greatly slowing down the compile).  This is likely
     to only be of interest to developers *of* the library.  It is
     used when running the selftest suite.  */
  GCC_JIT_BOOL_OPTION_SELFCHECK_GC,

  /* If true, gcc_jit_context_release will not clean up
     intermediate files written to the filesystem, and will display
     their location on stderr.  */
  GCC_JIT_BOOL_OPTION_KEEP_INTERMEDIATES,

  GCC_JIT_NUM_BOOL_OPTIONS
};

/* Set a string option on the given context.

   The context directly stores the (const char *), so the passed string
   must outlive the context.  */
extern void
gcc_jit_context_set_str_option (gcc_jit_context *ctxt,
				enum gcc_jit_str_option opt,
				const char *value);

/* Set an int option on the given context.  */
extern void
gcc_jit_context_set_int_option (gcc_jit_context *ctxt,
				enum gcc_jit_int_option opt,
				int value);

/* Set a boolean option on the given context.

   Zero is "false" (the default), non-zero is "true".  */
extern void
gcc_jit_context_set_bool_option (gcc_jit_context *ctxt,
				 enum gcc_jit_bool_option opt,
				 int value);

/* This actually calls into GCC and runs the build, all
   in a mutex for now.  The result is a wrapper around a .so file.
   It can only be called once on a given context.  */
extern gcc_jit_result *
gcc_jit_context_compile (gcc_jit_context *ctxt);

/* To be called after a compile, this gives the first error message
   that occurred on the context.

   The returned string is valid for the rest of the lifetime of the
   context.

   If no errors occurred, this will be NULL.  */
extern const char *
gcc_jit_context_get_first_error (gcc_jit_context *ctxt);

/* Locate a given function within the built machine code.
   This will need to be cast to a function pointer of the
   correct type before it can be called. */
extern void *
gcc_jit_result_get_code (gcc_jit_result *result,
			 const char *funcname);

/* Once we're done with the code, this unloads the built .so file.
   This cleans up the result; after calling this, it's no longer
   valid to use the result.  */
extern void
gcc_jit_result_release (gcc_jit_result *result);


/**********************************************************************
 Functions for creating "contextual" objects.

 All objects created by these functions share the lifetime of the context
 they are created within, and are automatically cleaned up for you when
 you call gcc_jit_context_release on the context.

 Note that this means you can't use references to them after you've
 released their context.

 All (const char *) string arguments passed to these functions are
 copied, so you don't need to keep them around.  Note that this *isn't*
 the case for other parts of the API.

 You create code by adding a sequence of statements to a function.
 Control flow is expressed using labels, rather than by explicitly
 managing blocks.   If you have blocks, you can simply create labels
 for the start of each block.  You can even create labels for every
 instruction in your language; unused labels will be optimized away.
**********************************************************************/

/**********************************************************************
 The base class of "contextual" object.
 **********************************************************************/
/* Which context is "obj" within?  */
extern gcc_jit_context *
gcc_jit_object_get_context (gcc_jit_object *obj);

/* Get a human-readable description of this object.
   The string buffer is created the first time this is called on a given
   object, and persists until the object's context is released.  */
extern const char *
gcc_jit_object_get_debug_string (gcc_jit_object *obj);

/**********************************************************************
 Debugging information.
 **********************************************************************/

/* Creating source code locations for use by the debugger.
   Line and column numbers are 1-based.  */
extern gcc_jit_location *
gcc_jit_context_new_location (gcc_jit_context *ctxt,
			      const char *filename,
			      int line,
			      int column);

/* Upcasting from location to object.  */
extern gcc_jit_object *
gcc_jit_location_as_object (gcc_jit_location *loc);


/**********************************************************************
 Types.
 **********************************************************************/

/* Upcasting from type to object.  */
extern gcc_jit_object *
gcc_jit_type_as_object (gcc_jit_type *type);

/* Access to specific types.  */
enum gcc_jit_types
{
  /* C's "void" type.  */
  GCC_JIT_TYPE_VOID,

  /* "void *".  */
  GCC_JIT_TYPE_VOID_PTR,

  /* Various integer types.  */

  /* C's "char" (of some signedness) and the variants where the
     signedness is specified.  */
  GCC_JIT_TYPE_CHAR,
  GCC_JIT_TYPE_SIGNED_CHAR,
  GCC_JIT_TYPE_UNSIGNED_CHAR,

  /* C's "short" and "unsigned short".  */
  GCC_JIT_TYPE_SHORT, /* signed */
  GCC_JIT_TYPE_UNSIGNED_SHORT,

  /* C's "int" and "unsigned int".  */
  GCC_JIT_TYPE_INT, /* signed */
  GCC_JIT_TYPE_UNSIGNED_INT,

  /* C's "long" and "unsigned long".  */
  GCC_JIT_TYPE_LONG, /* signed */
  GCC_JIT_TYPE_UNSIGNED_LONG,

  /* C99's "long long" and "unsigned long long".  */
  GCC_JIT_TYPE_LONG_LONG, /* signed */
  GCC_JIT_TYPE_UNSIGNED_LONG_LONG,

  /* Floating-point types  */

  GCC_JIT_TYPE_FLOAT,
  GCC_JIT_TYPE_DOUBLE,
  GCC_JIT_TYPE_LONG_DOUBLE,

  /* C type: (const char *).  */
  GCC_JIT_TYPE_CONST_CHAR_PTR,

 /* The C "size_t" type.  */
  GCC_JIT_TYPE_SIZE_T,

 /* C type: (FILE *)  */
  GCC_JIT_TYPE_FILE_PTR
};

extern gcc_jit_type *
gcc_jit_context_get_type (gcc_jit_context *ctxt,
			  enum gcc_jit_types type_);

/* Constructing new types. */

/* Given type "T", get type "T*".  */
extern gcc_jit_type *
gcc_jit_type_get_pointer (gcc_jit_type *type);

/* Given type "T", get type "const T".  */
extern gcc_jit_type *
gcc_jit_type_get_const (gcc_jit_type *type);

/* Struct-handling.  */
extern gcc_jit_field *
gcc_jit_context_new_field (gcc_jit_context *ctxt,
			   gcc_jit_location *loc,
			   gcc_jit_type *type,
			   const char *name);

/* Upcasting from field to object.  */
extern gcc_jit_object *
gcc_jit_field_as_object (gcc_jit_field *field);

extern gcc_jit_type *
gcc_jit_context_new_struct_type (gcc_jit_context *ctxt,
				 gcc_jit_location *loc,
				 const char *name,
				 int num_fields,
				 gcc_jit_field **fields);

/**********************************************************************
 Constructing functions.
 **********************************************************************/
extern gcc_jit_param *
gcc_jit_context_new_param (gcc_jit_context *ctxt,
			   gcc_jit_location *loc,
			   gcc_jit_type *type,
			   const char *name);

/* Upcasting from param to object.  */
extern gcc_jit_object *
gcc_jit_param_as_object (gcc_jit_param *param);

extern gcc_jit_lvalue *
gcc_jit_param_as_lvalue (gcc_jit_param *param);

extern gcc_jit_rvalue *
gcc_jit_param_as_rvalue (gcc_jit_param *param);

enum gcc_jit_function_kind
{
  /* Function is defined by the client code and visible
     by name outside of the JIT.  */
  GCC_JIT_FUNCTION_EXPORTED,

  /* Function is defined by the client code, but is invisible
     outside of the JIT.  Analogous to a "static" function.  */
  GCC_JIT_FUNCTION_INTERNAL,

  /* Function is not defined by the client code; we're merely
     referring to it.  Analogous to using an "extern" function from a
     header file.  */
  GCC_JIT_FUNCTION_IMPORTED

};


extern gcc_jit_function *
gcc_jit_context_new_function (gcc_jit_context *ctxt,
			      gcc_jit_location *loc,
			      enum gcc_jit_function_kind kind,
			      gcc_jit_type *return_type,
			      const char *name,
			      int num_params,
			      gcc_jit_param **params,
			      int is_variadic);

/* Upcasting from function to object.  */
extern gcc_jit_object *
gcc_jit_function_as_object (gcc_jit_function *func);

/* Create a label, to be placed later.

   The name can be NULL, or you can give it a meaningful name, which
   may show up in dumps of the internal representation.  */
extern gcc_jit_label *
gcc_jit_function_new_forward_label (gcc_jit_function *func,
				    const char *name);

/* Upcasting from label to object.  */
extern gcc_jit_object *
gcc_jit_label_as_object (gcc_jit_label *label);

/**********************************************************************
 lvalues, rvalues and expressions.
 **********************************************************************/

extern gcc_jit_lvalue *
gcc_jit_context_new_global (gcc_jit_context *ctxt,
			    gcc_jit_location *loc,
			    gcc_jit_type *type,
			    const char *name);

/* Upcasting.  */
extern gcc_jit_object *
gcc_jit_lvalue_as_object (gcc_jit_lvalue *lvalue);

extern gcc_jit_rvalue *
gcc_jit_lvalue_as_rvalue (gcc_jit_lvalue *lvalue);

extern gcc_jit_object *
gcc_jit_rvalue_as_object (gcc_jit_rvalue *rvalue);

/* Integer constants. */
extern gcc_jit_rvalue *
gcc_jit_context_new_rvalue_from_int (gcc_jit_context *ctxt,
				     gcc_jit_type *numeric_type,
				     int value);

extern gcc_jit_rvalue *
gcc_jit_context_zero (gcc_jit_context *ctxt,
		      gcc_jit_type *numeric_type);

extern gcc_jit_rvalue *
gcc_jit_context_one (gcc_jit_context *ctxt,
		     gcc_jit_type *numeric_type);

/* Floating-point constants.  */
extern gcc_jit_rvalue *
gcc_jit_context_new_rvalue_from_double (gcc_jit_context *ctxt,
					gcc_jit_type *numeric_type,
					double value);

/* Pointers.  */
extern gcc_jit_rvalue *
gcc_jit_context_new_rvalue_from_ptr (gcc_jit_context *ctxt,
				     gcc_jit_type *pointer_type,
				     void *value);

/* String literals. */
extern gcc_jit_rvalue *
gcc_jit_context_new_string_literal (gcc_jit_context *ctxt,
				    const char *value);

enum gcc_jit_unary_op
{
  /* Negate an arithmetic value; analogous to:
       -(EXPR)
     in C.  */
  GCC_JIT_UNARY_OP_MINUS,

  /* Bitwise negation of an integer value (one's complement); analogous
     to:
       ~(EXPR)
     in C.  */
  GCC_JIT_UNARY_OP_BITWISE_NEGATE,

  /* Logical negation of an arithmetic or pointer value; analogous to:
       !(EXPR)
     in C.  */
  GCC_JIT_UNARY_OP_LOGICAL_NEGATE
};

extern gcc_jit_rvalue *
gcc_jit_context_new_unary_op (gcc_jit_context *ctxt,
			      gcc_jit_location *loc,
			      enum gcc_jit_unary_op op,
			      gcc_jit_type *result_type,
			      gcc_jit_rvalue *rvalue);

enum gcc_jit_binary_op
{
  /* Addition of arithmetic values; analogous to:
       (EXPR_A) + (EXPR_B)
     in C.
     For pointer addition, use gcc_jit_context_new_array_lookup.  */
  GCC_JIT_BINARY_OP_PLUS,

  /* Subtraction of arithmetic values; analogous to:
       (EXPR_A) - (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_MINUS,

  /* Multiplication of a pair of arithmetic values; analogous to:
       (EXPR_A) * (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_MULT,

  /* Quotient of division of arithmetic values; analogous to:
       (EXPR_A) / (EXPR_B)
     in C.
     The result type affects the kind of division: if the result type is
     integer-based, then the result is truncated towards zero, whereas
     a floating-point result type indicates floating-point division.  */
  GCC_JIT_BINARY_OP_DIVIDE,

  /* Remainder of division of arithmetic values; analogous to:
       (EXPR_A) % (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_MODULO,

  /* Bitwise AND; analogous to:
       (EXPR_A) & (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_BITWISE_AND,

  /* Bitwise exclusive OR; analogous to:
       (EXPR_A) ^ (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_BITWISE_XOR,

  /* Bitwise inclusive OR; analogous to:
       (EXPR_A) | (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_BITWISE_OR,

  /* Logical AND; analogous to:
       (EXPR_A) && (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_LOGICAL_AND,

  /* Logical OR; analogous to:
       (EXPR_A) || (EXPR_B)
     in C.  */
  GCC_JIT_BINARY_OP_LOGICAL_OR
};

extern gcc_jit_rvalue *
gcc_jit_context_new_binary_op (gcc_jit_context *ctxt,
			       gcc_jit_location *loc,
			       enum gcc_jit_binary_op op,
			       gcc_jit_type *result_type,
			       gcc_jit_rvalue *a, gcc_jit_rvalue *b);

/* (Comparisons are treated as separate from "binary_op" to save
   you having to specify the result_type).  */

enum gcc_jit_comparison
{
  /* (EXPR_A) == (EXPR_B).  */
  GCC_JIT_COMPARISON_EQ,

  /* (EXPR_A) != (EXPR_B).  */
  GCC_JIT_COMPARISON_NE,

  /* (EXPR_A) < (EXPR_B).  */
  GCC_JIT_COMPARISON_LT,

  /* (EXPR_A) <=(EXPR_B).  */
  GCC_JIT_COMPARISON_LE,

  /* (EXPR_A) > (EXPR_B).  */
  GCC_JIT_COMPARISON_GT,

  /* (EXPR_A) >= (EXPR_B).  */
  GCC_JIT_COMPARISON_GE
};

extern gcc_jit_rvalue *
gcc_jit_context_new_comparison (gcc_jit_context *ctxt,
				gcc_jit_location *loc,
				enum gcc_jit_comparison op,
				gcc_jit_rvalue *a, gcc_jit_rvalue *b);

extern gcc_jit_rvalue *
gcc_jit_context_new_call (gcc_jit_context *ctxt,
			  gcc_jit_location *loc,
			  gcc_jit_function *func,
			  int numargs , gcc_jit_rvalue **args);

extern gcc_jit_rvalue *
gcc_jit_context_new_array_lookup (gcc_jit_context *ctxt,
				  gcc_jit_location *loc,
				  gcc_jit_rvalue *ptr,
				  gcc_jit_rvalue *index);

/* Field access is provided separately for both lvalues and rvalues.  */

/* Accessing a field of an lvalue of struct type, analogous to:
      (EXPR).field = ...;
   in C.  */
extern gcc_jit_lvalue *
gcc_jit_lvalue_access_field (gcc_jit_lvalue *struct_,
			     gcc_jit_location *loc,
			     gcc_jit_field *field);

/* Accessing a field of an rvalue of struct type, analogous to:
      (EXPR).field
   in C.  */
extern gcc_jit_rvalue *
gcc_jit_rvalue_access_field (gcc_jit_rvalue *struct_,
			     gcc_jit_location *loc,
			     gcc_jit_field *field);

/* Accessing a field of an rvalue of pointer type, analogous to:
      (EXPR)->field
   in C, itself equivalent to (*EXPR).FIELD  */
extern gcc_jit_lvalue *
gcc_jit_rvalue_dereference_field (gcc_jit_rvalue *ptr,
				  gcc_jit_location *loc,
				  gcc_jit_field *field);

/* Dereferencing a pointer; analogous to:
     *(EXPR)
*/
extern gcc_jit_lvalue *
gcc_jit_rvalue_dereference (gcc_jit_rvalue *rvalue,
			    gcc_jit_location *loc);

/* Taking the address of an lvalue; analogous to:
     &(EXPR)
   in C.  */
extern gcc_jit_rvalue *
gcc_jit_lvalue_get_address (gcc_jit_lvalue *lvalue,
			    gcc_jit_location *loc);

extern gcc_jit_lvalue *
gcc_jit_function_new_local (gcc_jit_function *func,
			    gcc_jit_location *loc,
			    gcc_jit_type *type,
			    const char *name);

/**********************************************************************
 Statement-creation.
 **********************************************************************/

/* Add evaluation of an rvalue, discarding the result
   (e.g. a function call that "returns" void).

   This is equivalent to this C code:

     (void)expression;
*/
extern void
gcc_jit_function_add_eval (gcc_jit_function *func,
			   gcc_jit_location *loc,
			   gcc_jit_rvalue *rvalue);

/* Add evaluation of an rvalue, assigning the result to the given
   lvalue.

   This is roughly equivalent to this C code:

     lvalue = rvalue;
*/
extern void
gcc_jit_function_add_assignment (gcc_jit_function *func,
				 gcc_jit_location *loc,
				 gcc_jit_lvalue *lvalue,
				 gcc_jit_rvalue *rvalue);

/* Add evaluation of an rvalue, using the result to modify an
   lvalue.

   This is analogous to "+=" and friends:

     lvalue += rvalue;
     lvalue *= rvalue;
     lvalue /= rvalue;
   etc  */
extern void
gcc_jit_function_add_assignment_op (gcc_jit_function *func,
				    gcc_jit_location *loc,
				    gcc_jit_lvalue *lvalue,
				    enum gcc_jit_binary_op op,
				    gcc_jit_rvalue *rvalue);

/* Add a no-op textual comment to the internal representation of the
   code.  It will be optimized away, but will be visible in the dumps
   seen via
     GCC_JIT_BOOL_OPTION_DUMP_INITIAL_TREE
   and
     GCC_JIT_BOOL_OPTION_DUMP_INITIAL_GIMPLE,
   and thus may be of use when debugging how your project's internal
   representation gets converted to the libgccjit IR.  */
extern void
gcc_jit_function_add_comment (gcc_jit_function *func,
			      gcc_jit_location *loc,
			      const char *text);

/* Add evaluation of an rvalue, branching on the result to the
   appropriate label.

   This is roughly equivalent to this C code:

     if (boolval)
       goto on_true;
     else
       goto on_false;

   on_true must be non-NULL.

   on_false may be NULL, in which case this is roughly equivalent to:

     if (boolval)
       goto on_true;
*/
extern void
gcc_jit_function_add_conditional (gcc_jit_function *func,
				  gcc_jit_location *loc,
				  gcc_jit_rvalue *boolval,
				  gcc_jit_label *on_true,
				  gcc_jit_label *on_false);

/* Add a label at the current location in the function.

   This is roughly equivalent to this C code:

      name:

   The name can be NULL, or you can give it a meaningful name, which
   may show up in dumps of the internal representation.  */
extern gcc_jit_label *
gcc_jit_function_add_label (gcc_jit_function *func,
			    gcc_jit_location *loc,
			    const char *name);

/* Place a pre-existing label at the current location in the function.

   This is roughly equivalent to this C code:

      name:

   The label must not have already been placed.  */
extern void
gcc_jit_function_place_forward_label (gcc_jit_function *func,
				      gcc_jit_location *loc,
				      gcc_jit_label *lab);

/* Add a jump from the current location in the function to the given
   label.

   This is roughly equivalent to this C code:

      goto target;
*/
extern void
gcc_jit_function_add_jump (gcc_jit_function *func,
			   gcc_jit_location *loc,
			   gcc_jit_label *target);

/* Add evaluation of an rvalue, returning the value.

   This is roughly equivalent to this C code:

      return expression;
*/
extern void
gcc_jit_function_add_return (gcc_jit_function *func,
			     gcc_jit_location *loc,
			     gcc_jit_rvalue *rvalue);

/* Helper function for creating a loop:
      while (boolval)
	{
	   BODY;
	}

  Statements will be added to the body of the loop until
  gcc_jit_loop_end is called.  */
extern gcc_jit_loop *
gcc_jit_function_new_loop (gcc_jit_function *func,
			   gcc_jit_location *loc,
			   gcc_jit_rvalue *boolval);

/* Upcasting from loop to object.  */
extern gcc_jit_object *
gcc_jit_loop_as_object (gcc_jit_loop *loop);

extern void
gcc_jit_loop_end (gcc_jit_loop *loop,
		  gcc_jit_location *loc);

/**********************************************************************
 Nested contexts.
 **********************************************************************/

/* Given an existing JIT context, create a child context.

   The child inherits a copy of all option-settings from the parent.

   The child can reference objects created within the parent, but not
   vice-versa.

   The lifetime of the child context must be bounded by that of the
   parent: you should release a child context before releasing the parent
   context.

   If you use a function from a parent context within a child context,
   you have to compile the parent context before you can compile the
   child context, and the gcc_jit_result of the parent context must
   outlive the gcc_jit_result of the child context.

   This allows caching of shared initializations.  For example, you could
   create types and declarations of global functions in a parent context
   once within a process, and then create child contexts whenever a
   function or loop becomes hot. Each such child context can be used for
   JIT-compiling just one function or loop, but can reference types
   and helper functions created within the parent context.

   Contexts can be arbitrarily nested, provided the above rules are
   followed, but it's probably not worth going above 2 or 3 levels, and
   there will likely be a performance hit for such nesting.  */

extern gcc_jit_context *
gcc_jit_context_new_child_context (gcc_jit_context *parent_ctxt);

#ifdef __cplusplus
}
#endif /* __cplusplus */

#endif  /* LIBGCCJIT_H  */
