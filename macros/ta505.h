#ifndef INDLL_H
#define INDLL_H

   #ifdef EXPORTING_DLL
      extern __declspec(dllexport) void main() ;
   #else
      extern __declspec(dllimport) void main() ;
   #endif

#endif