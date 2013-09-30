// BrowseForFolerModule.h
//

#if !defined(BrowseForFolerModule_h)
#define BrowseForFolerModule_h

#include "resource.h"


class BrowseForFolerModule
{
public:

    //static void CoTaskMemFree(Long ByVal pv);
    //static void CopyMemory(Any pDest, Any pSource, Long ByVal dwLength);
    //static Long LocalAlloc(Long ByVal uFlags, Long ByVal uBytes);
    //static Long LocalFree(Long ByVal hMem);
    //static Long BrowseCallbackProcStr(Long ByVal hwnd, Long ByVal uMsg, Long ByVal lParam, Long ByVal lpData);
    //static Long FunctionPointer(Long FunctionAddress);
    //static String BrowseForFolder(Long hwnd, String selectedPath);
    static void BrowseForFoldersFromPathStart(HWND hwnd,CString PathStart, CString &PathSelected);
    //static int	WM_USER;
    //static Long	BFFM_SETSELECTIONA;
    //static Long	BFFM_SETSELECTIONW;
    //static int	LPTR;

protected:


private:


};

#endif /* BrowseForFolerModule_h */
