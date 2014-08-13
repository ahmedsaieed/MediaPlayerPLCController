Attribute VB_Name = "DMT"

'/////////////////////////////////////////////////////////////////////////
'//
'//  Name:
'//      DMT.bas
'//
'//  Description:
'//      DMT Library VB6 function declaration module file
'//
'//  History:
'//      Date            Author          Plant           Version         Comment
'//      13/02/2012      AllenCX         Taoyuan III     Version 2.2
'//      05/01/2009      Anderson        Taoyuan I       Version 2.0
'//      10/01/2007      Anderson        Taoyuan I       Version 1.1
'//      08/01/2007      Anderson        Taoyuan I       Version 1.0
'//
'/////////////////////////////////////////////////////////////////////////

'// Data Access
Declare Function RequestData Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal func_code As Long, ByRef sendbuf As Byte, ByVal sendlen As Long) As Long
Declare Function ResponseData Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByRef slave_addr As Long, ByRef func_code As Long, ByRef recvbuf As Byte) As Long

'// Serial Communication
Declare Function OpenModbusSerial Lib "DMT.dll" (ByVal conn_num As Long, ByVal baud_rate As Long, ByVal data_len As Long, ByVal parity As Byte, ByVal stop_bits As Long, ByVal modbus_mode As Long) As Long
Declare Sub CloseSerial Lib "DMT.dll" (ByVal conn_num As Long)
Declare Function GetLastSerialErr Lib "DMT.dll" () As Long
Declare Sub ResetSerialErr Lib "DMT.dll" ()

'// Socket Communication
Declare Function OpenModbusTCPSocket Lib "DMT.dll" (ByVal conn_num As Long, ByVal ipaddr As Long) As Long
Declare Sub CloseSocket Lib "DMT.dll" (ByVal conn_num As Long)
Declare Function GetLastSocketErr Lib "DMT.dll" () As Long
Declare Sub ResetSocketErr Lib "DMT.dll" ()
Declare Function ReadSelect Lib "DMT.dll" (ByVal conn_num As Long, ByVal millisecs As Long) As Long

'// MODBUS Address Calculation
Declare Function DevToAddrA Lib "DMT.dll" (ByVal series As String, ByVal device As String, ByVal qty As Long) As Long
'// Wrapped MODBUS Funcion : 0x01
Declare Function ReadCoilsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_r As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x02
Declare Function ReadInputsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_r As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x03
Declare Function ReadHoldRegsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_r As Long, ByVal req As String, ByVal res As String) As Long
Declare Function ReadHoldRegs32A Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_r As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x04
Declare Function ReadInputRegsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_r As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x05
Declare Function WriteSingleCoilA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal data_w As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x06
Declare Function WriteSingleRegA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal data_w As Long, ByVal req As String, ByVal res As String) As Long
Declare Function WriteSingleReg32A Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal data_w As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x0F
Declare Function WriteMultiCoilsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_w As Long, ByVal req As String, ByVal res As String) As Long
'// Wrapped MODBUS Funcion : 0x10
Declare Function WriteMultiRegsA Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_w As Long, ByVal req As String, ByVal res As String) As Long
Declare Function WriteMultiRegs32A Lib "DMT.dll" (ByVal comm_type As Long, ByVal conn_num As Long, ByVal slave_addr As Long, ByVal dev_addr As Long, ByVal qty As Long, ByRef data_w As Long, ByVal req As String, ByVal res As String) As Long

