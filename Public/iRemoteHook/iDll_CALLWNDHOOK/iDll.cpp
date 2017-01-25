#include "stdafx.h"
#include "iDll.h"
#include <windows.h>
#include <stdlib.h>

//注意：VC++ Dll的导出函数若要给VB调用，需要__stdcall的调用方式，extren "C"修饰，和def文件

extern "C"
{

#pragma data_seg("RemoteHookDllData")//设置共享数据段

HHOOK hHook=NULL;//钩子句柄

#define Msg_list_MaxSize 10
int Msg_list[Msg_list_MaxSize+1]={0};//需传输的消息集合，Msg_list[0]存储消息个数
#define hwnd_list_MaxSize 5
int hwnd_list[Msg_list_MaxSize+1][hwnd_list_MaxSize+1]={{0}};//需传输消息对应的接收句柄集合，hwnd_list[i][0]存储第i个消息的接收句柄数

//消息注册个数上限为10个，对于一个消息接收句柄上限为5个，这是为了保证程序运行的高效

#pragma data_seg()
	
HINSTANCE hInsDll;//Dll实例句柄
	
LRESULT CALLBACK iProc(int nCode,WPARAM wParam,LPARAM lParam);//Hook回调函数
	

DLLIMPORT bool __stdcall RegisterMessage(int Msg,int hReciver)//注册消息
{
	int i;
	for(i=1;i<=Msg_list[0];i++) if(Msg_list[i]==Msg) break;//查找消息是否已注册
	if(i>Msg_list[0])//如果消息未注册，则注册
	{
		if(Msg_list[0]>=Msg_list_MaxSize)
		{
			MessageBox(0,"Overflow(Msg_list)!","iRemoteHook - Dll",MB_ICONINFORMATION);
			return false;
		}
		Msg_list[++Msg_list[0]]=Msg;
		hwnd_list[Msg_list[0]][0]=1;
		hwnd_list[Msg_list[0]][1]=hReciver;
	}
	else//如果消息已注册，则增加接收句柄
	{
		if(hwnd_list[i][0]>=hwnd_list_MaxSize)
		{
			MessageBox(0,"Overflow(hwnd_list)!","iRemoteHook - Dll",MB_ICONINFORMATION);
			return false;
		}
		hwnd_list[i][++hwnd_list[i][0]]=hReciver;
	}
	return true;
}

DLLIMPORT int __stdcall SetHook(int TID)//挂钩
{
	if(hHook!=NULL) return (int)hHook;
	hHook=SetWindowsHookEx(WH_CALLWNDPROC,iProc,hInsDll,(DWORD)TID);
	if(!hHook)
	{
		char S[25];
		itoa((int)GetLastError(),S,10);
		MessageBox(0,strcat("Fail to set hook! GetLastError = ",S),"iRemoteHook - Dll",MB_ICONINFORMATION);
	}
	return (int)hHook;
}
	
LRESULT CALLBACK iProc(int nCode,WPARAM wParam,LPARAM lParam)//Hook回调函数
{
#define CWP ((CWPSTRUCT*)lParam)
	if(nCode==HC_ACTION)
	{
		for(int i=1;i<=Msg_list[0];i++)
		{
			if(CWP->message==Msg_list[i])
			{
				int Datas[4]={(int)CWP->lParam,(int)CWP->wParam,(int)CWP->message,(int)CWP->hwnd};
				//这里需要额外分配新的空间，不能将lParam直接作为COPYDATASTRUCT的参数发送

				COPYDATASTRUCT cds;
				cds.dwData=0;
				cds.cbData=sizeof(Datas);
				cds.lpData=Datas;

				for(int j=1;j<=hwnd_list[i][0];j++)
					SendMessageA((HWND)hwnd_list[i][j],WM_COPYDATA,0,(LPARAM)&cds);//发送WM_COPYDATA传回消息
				//WM_COPYDATA可以用于跨进程传输消息
				//WH_CALLWNDPROC不能修改消息，故没有根据SendMessageA的返回值作处理
			}
		}
			
	}
		
	return CallNextHookEx(hHook,nCode,wParam,lParam);
}

DLLIMPORT void __stdcall UnHook()
{
	UnhookWindowsHookEx(hHook);
}
}

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
    switch (ul_reason_for_call)
	{
		case DLL_PROCESS_ATTACH:
			hInsDll=(HINSTANCE)hModule;//保存Dll实例句柄
		case DLL_THREAD_ATTACH:
		case DLL_THREAD_DETACH:
		case DLL_PROCESS_DETACH:
			break;
    }
    return TRUE;
}
