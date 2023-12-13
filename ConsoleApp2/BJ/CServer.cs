using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace BJ
{
    public class CServer
    {
        //CServer_CTEC_D.dll
        public const string dll_name = "CServer_CTEC_D.dll";  //CServer_CTEC_D.dll  CClient_CTEC_D.dll
        public CServer()
        {
        }

        /// <summary>
        /// 初始化服务器连接 , CallingConvention = CallingConvention.Cdecl
        /// </summary>
        /// <returns></returns>
        [DllImport(dll_name, EntryPoint = "InitServer", CallingConvention = CallingConvention.Cdecl)]
        public static extern int InitServer(string ipadrr, int uPort);

        /// </summary>
        /// 获取数据长度
        /// <param name="uSoserport></param>
        /// <param name="iTimeout"></param>
        /// <returns></returns>
        [DllImport(dll_name, EntryPoint = "GetDataLen", CallingConvention = CallingConvention.Cdecl)]
        public static extern int GetDataLen(ref int uSocket, int iTimeout);
        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="revdata"></param>
     [DllImport(dll_name, EntryPoint = "GetDataLen", CallingConvention = CallingConvention.Cdecl)]
        public static extern int GetDataLen1(int uSocket, int iTimeout);        /// <param name="datalen"></param>
        /// <param name="usocket"></param>
        /// <returns></returns>
        [DllImport(dll_name, EntryPoint = "RecvData", CallingConvention = CallingConvention.Cdecl)]
        public static extern int RecvData(byte[] revData, int datalen, int usocket);
        /// <summary>
        /// 检查服务器连接
        /// </summary>
        /// <returns></returns>
        [DllImport(dll_name, EntryPoint = "GetServerState", CallingConvention = CallingConvention.Cdecl)]
        public static extern int GetServerState();
        /// <summary>
        /// 关闭服务
        /// </summary>
        /// <returns></returns>
        [DllImport(dll_name, EntryPoint = "Exit_Server", CallingConvention = CallingConvention.Cdecl)]
        public static extern int Exit_Server(int uSocket);
    }
}
