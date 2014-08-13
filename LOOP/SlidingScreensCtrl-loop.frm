VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sliding Screens Controller"
   ClientHeight    =   705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Label Label1 
      Caption         =   "Initializing system..."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    strDev = "M17"
    addr = DevToAddrA(strProduct, strDev, dev_qty)
    Call OpenModbusSerial(conn_num, 115200, 7, AscB("E"), 1, modbus_mode)
    ret = WriteSingleCoilA(comm_type, conn_num, slave_addr, addr, data_to_dev, req_s, res_s)
    Call CloseSerial(conn_num)
    Label1.Caption = "Executing program..."
    End
End Sub
