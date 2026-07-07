# AGENTS.md

This repository is the **GNU Compiler Collection (GCC)** source tree. The "application" is the
compiler toolchain itself (`gcc`, `g++`, etc.), built from source.

## Cursor Cloud specific instructions

### Environment overview
- Host toolchain: system `gcc`/`g++` 13.3 (Ubuntu), `make`, `flex`, `bison`, `dejagnu` are installed.
- Build prerequisites (`libgmp-dev`, `libmpfr-dev`, `libmpc-dev`, `libisl-dev`, `libzstd-dev`,
  `flex`, `bison`, `texinfo`) are installed via `apt` by the startup update script. GCC's
  `contrib/download_prerequisites` is therefore NOT needed.

### Building (do NOT build in the source tree)
- Always configure/build in a **separate build directory**. The one used during setup is
  `~/gcc-build` (installed to `~/gcc-install`). Recreate with:
  ```
  mkdir -p ~/gcc-build && cd ~/gcc-build
  /workspace/configure --enable-languages=c,c++ --disable-bootstrap --disable-multilib --prefix="$HOME/gcc-install"
  make -j$(nproc)
  ```
- `--disable-bootstrap` builds the compiler once (bootstrap builds it 3x). A full C/C++ build takes
  ~9 min on 4 cores. Add languages via `--enable-languages=...` (e.g. `c,c++,fortran`) only if needed
  — each language significantly increases build time.
- `make install` (fast, ~10s) populates `~/gcc-install`. Use `~/gcc-install/bin/gcc` and
  `~/gcc-install/bin/g++` for a fully functional installed compiler. When running installed C++
  binaries, set `LD_LIBRARY_PATH=~/gcc-install/lib64` (or install adds an rpath if configured).
- The in-tree driver is `~/gcc-build/gcc/xgcc -B~/gcc-build/gcc/` but it is fiddly for C++ (libstdc++
  headers not on the default path); prefer the installed compiler for C++ smoke tests.
- After editing compiler sources, rerun `make -j$(nproc)` in `~/gcc-build` (incremental) then
  `make install`. There is no hot-reload.

### Testing (DejaGnu based)
- `make check` runs the full testsuite and takes **hours**. Scope it with `RUNTESTFLAGS`.
- Run C tests from `~/gcc-build`: `make check-gcc RUNTESTFLAGS="dg.exp=<glob>.c"`
  (e.g. `dg.exp=c99-*.c`). Results land in `gcc/testsuite/gcc/gcc.{sum,log}`.
- Run C++ tests from the **`~/gcc-build/gcc` subdirectory** (not the top build dir):
  `make check-g++ RUNTESTFLAGS="dg.exp=<glob>.C"`. Results land in `gcc/testsuite/g++/g++.{sum,log}`.
- The `RUNTESTFLAGS` filter must be a single space-free token; use one glob (e.g. `c99-*.c`), not
  space-separated lists.

### Lint / style
- GCC has no traditional linter. Coding-style checks for patches use
  `contrib/check_GNU_style.sh <patch>` (and `contrib/check_GNU_style_lib.sh`). Warnings during the
  build (e.g. `-Wformat`) are expected in the current tree and are not build failures.
