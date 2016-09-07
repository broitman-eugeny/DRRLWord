#include "C:\Program Files\Mathsoft\Mathcad 11 Enterprise Edition\UserEFI\Microsft\include\mcadincl.h"


extern FUNCTIONINFO Ftest_WORD;

    char *ErrorMessageTable[] = {
    "argument must be real",    //  error 1 --  argument must be real
    "insufficient memory",      //  error 2 --  memory allocation error
    "interrupted",               //  error 3 --  execution interrupted
    "Не выбран файл",     //  error 4 --  Не выбран файл профиля
    "Отчет не выведен"  //error 5 --  Отрицательный просвет при средней рефракции в точке критического препятствия
    };




HINSTANCE g_hInst;		// the current instance

#pragma argsused

BOOL WINAPI DllEntryPoint (HANDLE hDLL, DWORD dwReason, LPVOID lpReserved)
{

   // Save off the handle to the current instance.
    g_hInst = (HINSTANCE)hDLL;

  switch (dwReason)
  {
    case DLL_PROCESS_ATTACH:
    {

      // DLL is attaching to the address space of the current process.
      //
		  if (!CreateUserErrorMessageTable( (HINSTANCE)hDLL, 4, ErrorMessageTable ) )
        {
				MessageBox(NULL,	// handle of owner window
    						"ErrorMessageTable - NULL",	// address of text in message box
    						"CreateUserErrorMessageTable",	// address of title of message box
    						MB_ICONASTERISK); 	// style of message box
            break;
        }


 		  if ( CreateUserFunction( (HINSTANCE)hDLL, &Ftest_WORD) == NULL )
        {
        		MessageBox(NULL,	// handle of owner window
    						"Ftest_WORD - NULL",	// address of text in message box
    						"CreateUserFunction",	// address of title of message box
    						MB_ICONASTERISK); 	// style of message box
            break;
        }
        
    }

	 case DLL_THREAD_ATTACH:        // A new thread is being created in the current process.
	 case DLL_THREAD_DETACH:        // A thread is exiting cleanly.
	 case DLL_PROCESS_DETACH:      // The calling process is detaching the DLL from its address space.

		break;
  }
  return TRUE;
}//DllEntryPoint

