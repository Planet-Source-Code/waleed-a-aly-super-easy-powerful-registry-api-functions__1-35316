VERSION 5.00
Begin VB.Form frmRegistryAPI 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registry"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   Icon            =   "frmRegistryAPI.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1515
   End
   Begin VB.CommandButton cmdReg 
      Caption         =   "Perform Registry Operations"
      Height          =   795
      Left            =   1500
      TabIndex        =   0
      Top             =   540
      Width           =   1515
   End
End
Attribute VB_Name = "frmRegistryAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'****************************************************************************'
'  Currently these API Registry functions support [Setting] and [Extracting] '
' String & Binary Values Only -since they are the most common to use- but    '
' still, it is easy to support the rest formats.                             '
'  If you add such support to my functions please eMail them to me, I do not '
' currently intend to work on supporting the rest formats 'coz i don't need  '
' them in my programs.                                                       '
'****************************************************************************'

Private Sub cmdReg_Click()
'In general, functions dealing with [Keys] require the [Full Path] to the key
'while functions dealing with [Values] require the [Key Handle] for the containing key
'note that the [Key Handle] is changed if Key is [ReOpened] or [ReCreated]

    'Create a new Registry Key
    Debug.Print CreateRegKey(HKEY_LOCAL_MACHINE, StartUpEntry + "\My Entry", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS)
    'Place its Handle number on the Button
    cmdReg.Caption = GetKeyHandle
    
    'the following step is useless, it's just provided here for testing the function
    'it is useless because the key handle is already returned by the [CreateRegKey] function
    Debug.Print OpenRegKey(HKEY_LOCAL_MACHINE, StartUpEntry + "\My Entry", KEY_ALL_ACCESS)
    'Place the New Handle (Key is ReOpened) in the text box
    txt = GetKeyHandle
    
    'try setting the maximum & the minimum Binary Values (Long Number) [4Bytes]
    Debug.Print SetRegValue(GetKeyHandle, "My Binary Box", REG_BINARY, "+2147483647")
    Debug.Print QueryRegValue(GetKeyHandle, "My Binary Box")
    'note that the Numerical Value was passed as a sring, it can still be passed as a
    'Numerical Value or a Variable as follows (Passed Data is of Type Variant)
    Debug.Print SetRegValue(GetKeyHandle, "My Binary Box", REG_BINARY, -2147483647)
    Debug.Print QueryRegValue(GetKeyHandle, "My Binary Box")
    'now delete the Binary Box
    Debug.Print DeleteRegValue(GetKeyHandle, "My Binary Box")
    
    'try setting a String Value
    Debug.Print SetRegValue(GetKeyHandle, "My xBox", REG_SZ, "Hola! :D 12345")
    Debug.Print QueryRegValue(GetKeyHandle, "My xBox")
    'now delete the xBox
    Debug.Print DeleteRegValue(GetKeyHandle, "My xBox")

    'now delete the whole Key & Close its Handle
    Debug.Print DeleteRegKey(HKEY_LOCAL_MACHINE, StartUpEntry + "\My Entry")
    Debug.Print CloseRegKey(GetKeyHandle)

'note:
'for floating-point Numbers or Integers larger than the range of Long,
'you cannot save them as Binary data, instead save them as strings.
End Sub
