VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSpawner 
   Appearance      =   0  'Flat
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Item Spawner - Equipment -> 47 items"
   ClientHeight    =   4935
   ClientLeft      =   8280
   ClientTop       =   4425
   ClientWidth     =   9015
   Icon            =   "frmItemSpawner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList itemsImageList 
      Left            =   7710
      Top             =   4290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CheckBox chkClose 
      Caption         =   "Close window after spawning"
      ForeColor       =   &H00C0C000&
      Height          =   420
      Left            =   4590
      TabIndex        =   10
      Top             =   285
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin MSComctlLib.ListView listItems 
      Height          =   3720
      Left            =   45
      TabIndex        =   9
      Top             =   1185
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   6562
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      HotTracking     =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSpawn 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      Caption         =   "Spawn it"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3540
      TabIndex        =   8
      Top             =   345
      Width           =   795
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2400
      TabIndex        =   7
      Text            =   "1"
      Top             =   300
      Width           =   705
   End
   Begin VB.OptionButton radioInv 
      Caption         =   "Inventory(3 slots)"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   540
      Value           =   -1  'True
      Width           =   2340
   End
   Begin VB.OptionButton radioGround 
      Caption         =   "Ground"
      Enabled         =   0   'False
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
   Begin MSComctlLib.TabStrip tabItems 
      Height          =   4185
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   7382
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   11
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recent"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "None"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Equipment"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Consumable"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Title"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Spell"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Teleport"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset Stats"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto Life"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sprite Change"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Recipe"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2625
      ScaleHeight     =   270
      ScaleWidth      =   2685
      TabIndex        =   12
      Top             =   2610
      Width           =   2685
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Caption         =   "No items available in this category!"
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   15
         TabIndex        =   13
         Top             =   30
         Width           =   2655
      End
   End
   Begin VB.Label lblMax 
      Caption         =   "/max"
      Height          =   255
      Left            =   3120
      TabIndex        =   15
      Top             =   375
      Width           =   405
   End
   Begin VB.Label lblHelp2 
      BackStyle       =   0  'Transparent
      Caption         =   "and ""Spawn It""."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   6840
      TabIndex        =   14
      Top             =   480
      Width           =   1305
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOptions 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   4950
      TabIndex        =   11
      Top             =   15
      Width           =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   3
      X1              =   4695
      X2              =   5855
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Line lineAmount 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      X1              =   2670
      X2              =   4230
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Label lblAmount 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   195
      Left            =   3150
      TabIndex        =   6
      Top             =   30
      Width           =   645
   End
   Begin VB.Label lblHelp1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose the item, input Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   465
      Left            =   6090
      TabIndex        =   5
      Top             =   270
      Width           =   2805
      WordWrap        =   -1  'True
   End
   Begin VB.Line lineHow 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   6120
      X2              =   8280
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lblHow 
      Alignment       =   1  'Right Justify
      Caption         =   "How to use it"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   6720
      TabIndex        =   4
      Top             =   0
      Width           =   1185
   End
   Begin VB.Line lineWhere 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      X1              =   0
      X2              =   2160
      Y1              =   210
      Y2              =   210
   End
   Begin VB.Label lblWhere 
      Caption         =   "Where to Spawn It"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   1665
   End
End
Attribute VB_Name = "frmItemSpawner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lastTab As Integer
Private allowTitle As Boolean
Private currentItemIndex As Integer
Private currentAmount As Long
Private picked As Boolean
Private freeInvSlots As Byte
Private currentMaxLimit As Long
Private descIndex As Long
Public updatingItem As Boolean

Private Declare Function SendMessage Lib "user32" Alias _
 "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, lParam As Any) As Long
 
Public Function ListView_SetIconSpacing(hWndLV As Long, cx As Long, cy As Long) As Long
    Dim LVM_SETICONSPACING As Long
    
    LVM_SETICONSPACING = 4149
    ListView_SetIconSpacing = SendMessage(hWndLV, LVM_SETICONSPACING, 0, ByVal MakeLong(cx, cy))
End Function

Public Sub updateFreeSlots()

    If Me.Visible Then
        freeInvSlots = countFreeSlots
        If freeInvSlots = 0 Then
            radioInv.Enabled = False
            radioInv.Caption = "Inventory(No slots available)"
            radioGround.Enabled = True
            radioGround.Value = True
        Else
            radioInv.Caption = "Inventory(" & freeInvSlots & " free slot" & IIf(freeInvSlots > 1, "s", "") & ")"
            radioInv.Enabled = True
            radioGround.Enabled = True
        End If
        updateMaxLimit
    End If
    
End Sub

Private Sub updateMaxLimit()
    If Me.Visible = True And listItems.listItems.Count > 0 Then
        Dim stackable As Boolean
        stackable = Item(currentlyListedIndexes(listItems.SelectedItem.Index - 1)).stackable
        If radioGround.Value And Not stackable Then
            currentMaxLimit = MAX_MAP_ITEMS
        ElseIf Not stackable Then
            currentMaxLimit = freeInvSlots
        Else
            currentMaxLimit = 2147483468
            lblMax.Caption = "/---"
            Exit Sub
        End If
        lblMax.Caption = "/" & currentMaxLimit
    Else
            lblMax.Caption = "/NaN"
    End If
End Sub

Private Sub styleListwView(sType As Status, Optional Msg As String = vbNullString)
    Select Case sType
        Case Status.Correct
            listItems.BackColor = &H8000000E
            listItems.BorderStyle = ccFixedSingle
            picInfo.Visible = False
            picInfo.ZOrder 1
        Case Status.Error
            listItems.BackColor = &H8000000F
            listItems.BorderStyle = ccNone
            picInfo.Visible = True
            picInfo.ZOrder 0
            lblInfo.Caption = Msg
    End Select
End Sub

Private Function generateItemsForTab(tabNum As Byte) As Boolean
    Dim I As Long, Z As Long, tempItems() As ItemRec, ret As Boolean
    
    tabNum = tabNum - 2
    
    Select Case tabNum
        Case ITEM_TYPE_NONE
            ret = populateSpecificType(tempItems, ITEM_TYPE_NONE)
        Case ITEM_TYPE_EQUIPMENT
            ret = populateSpecificType(tempItems, ITEM_TYPE_EQUIPMENT)
        Case ITEM_TYPE_CONSUME
            ret = populateSpecificType(tempItems, ITEM_TYPE_CONSUME)
        Case ITEM_TYPE_TITLE
            ret = populateSpecificType(tempItems, ITEM_TYPE_TITLE)
        Case ITEM_TYPE_SPELL
            ret = populateSpecificType(tempItems, ITEM_TYPE_SPELL)
        Case ITEM_TYPE_TELEPORT
            ret = populateSpecificType(tempItems, ITEM_TYPE_TELEPORT)
        Case ITEM_TYPE_RESETSTATS
            ret = populateSpecificType(tempItems, ITEM_TYPE_RESETSTATS)
        Case ITEM_TYPE_AUTOLIFE
            ret = populateSpecificType(tempItems, ITEM_TYPE_AUTOLIFE)
        Case ITEM_TYPE_SPRITE
            ret = populateSpecificType(tempItems, ITEM_TYPE_SPRITE)
        Case ITEM_TYPE_RECIPE
            ret = populateSpecificType(tempItems, ITEM_TYPE_RECIPE)
    End Select
    
    If ret Then
        Set listItems.Icons = itemsImageList
                
        For I = 0 To UBound(tempItems)
            listItems.listItems.Add , , Trim$(tempItems(I).Name), itemsImageList.ListImages(I + 1).Index
        Next
        currentItemIndex = 0
        generateItemsForTab = True
    End If
End Function

Private Sub generateRecentItems()
Dim I As Byte
    If ArrayIsInitialized(lastSpawnedItems) = 0 Then
        styleListwView Status.Error, "You haven't spawned any items yet!"
    Else
        styleListwView Status.Correct
        For I = 0 To UBound(lastSpawnedItems) - 1
            itemsImageList.ListImages.Add , , LoadPictureGDIPlus(App.Path & GFX_PATH & "items\" & Item(lastSpawnedItems(I)).Pic & GFX_EXT, False, 32, 32, 16777215)
        Next I
        Set listItems.Icons = itemsImageList
                
        For I = 0 To UBound(lastSpawnedItems) - 1
            listItems.listItems.Add , , Trim$(Item(lastSpawnedItems(I)).Name), itemsImageList.ListImages(I + 1).Index
        Next
        cmdSpawn.Enabled = True
        currentItemIndex = 0
    End If
End Sub

Private Sub cmdSpawn_Click()
    Dim Item As Byte
    Dim I As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If tabItems.SelectedItem.Index > 1 Then
        Item = currentlyListedIndexes(currentItemIndex)
    Else
        Item = lastSpawnedItems(currentItemIndex)
    End If

    SendSpawnItem Item, CLng(txtAmount), IIf(radioGround.Value, True, False)

    Dim found As Integer, limit As Integer

    found = -1
    If ArrayIsInitialized(lastSpawnedItems) Then
     If UBound(lastSpawnedItems) > 0 Then
        For I = 0 To UBound(lastSpawnedItems) - 1
            If lastSpawnedItems(I) = Item Then
                found = I
                Exit For
            End If
        Next
    End If
    Else
        ReDim lastSpawnedItems(0) As Byte
    End If
    
    If found = -1 Then
        If UBound(lastSpawnedItems) = 20 Then
            DeleteByPtr lastSpawnedItems, 20
        End If
        InsertByPtr lastSpawnedItems, 0
    Else
        DeleteByPtr lastSpawnedItems, found
        InsertByPtr lastSpawnedItems, 0
    End If

    lastSpawnedItems(0) = Item
    frmAdmin.UpdateRecentSpawner
    
    If chkClose.Value = 1 Then
        Unload Me
        frmAdmin.lastIndex = -1
        frmAdmin.currentCategory = "Categories"
        lastTab = -1
        currentItemIndex = -1
    Else
        updatingItem = True
        tabItems_Click
    End If
    
    ' Error handler
ErrorHandler:
    HandleError "cmdSpawn_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    ListView_SetIconSpacing listItems.hWnd, 105, 56
    Move frmAdmin.Left - Width, frmAdmin.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If FormVisible("frmAdmin") Then
        frmAdmin.styleButtons
        frmAdmin.lastIndex = -1
        frmAdmin.currentCategory = "Categories"
        frmAdmin.picSpawner.Visible = False
    End If

    lastTab = -1
    currentItemIndex = -1
End Sub

Private Sub listItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdSpawn.Enabled = True
    currentItemIndex = Item.Index - 1
    picked = True
    updateMaxLimit
    Me.Caption = "Item Spawner - Going to spawn " & txtAmount.text & " " & listItems.listItems(Item.Index).text
End Sub

Private Sub listItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oListItem As ListItem, indexx As Long, num As Long
    
    Set oListItem = listItems.HitTest(X, Y)
    If Not oListItem Is Nothing Then
        indexx = oListItem.Index
        If Not FormVisible("frmItemDesc") Then
            Load frmItemDesc
        End If
        If indexx <> descIndex Then
            
            If tabItems.SelectedItem.Index = 1 Then
                num = lastSpawnedItems(indexx - 1)
            Else
                num = currentlyListedIndexes(indexx - 1)
            End If
            frmItemDesc.lblName = Trim$(Item(num).Name)
            frmItemDesc.lblStack = "Stackable: " & IIf(Item(num).stackable > 0, "yes", "no")
            frmItemDesc.lblLevel = "LVL: " & Item(num).LevelReq
            frmItemDesc.lblType = "Type: " & getItemType(Item(num).Type)
            frmItemDesc.lbl2Hand = "2-Handed: " & IIf(Item(num).TwoHanded > 0, "yes", "no")
            
            frmItemDesc.Visible = True
        End If
        descIndex = indexx
    Else
        If FormVisible("frmItemDesc") Then
            Unload frmItemDesc
        End If
        descIndex = 0
    End If
End Sub

Private Sub radioGround_Click()
    updateMaxLimit
End Sub

Private Sub radioInv_Click()
    updateMaxLimit
End Sub

Public Sub tabItems_Click()
    If lastTab = tabItems.SelectedItem.Index And Not updatingItem Then Exit Sub
    
    cmdSpawn.Enabled = False
    listItems.listItems.Clear
    Set listItems.Icons = Nothing
    Set listItems.SmallIcons = Nothing
    itemsImageList.ListImages.Clear
    
    If tabItems.SelectedItem.Index = 1 Then
        generateRecentItems
    Else
        If generateItemsForTab(tabItems.SelectedItem.Index) Then
            styleListwView Status.Correct
            cmdSpawn.Enabled = True
            picked = True
        Else
            styleListwView Status.Error, "No items available in this category!"
            picked = False
        End If
    End If
    
    updateMaxLimit
    
    If updatingItem Then
        updatingItem = False
        Exit Sub
    End If
    
    If frmAdmin.ignoreChange Then
        frmAdmin.ignoreChange = False
    ElseIf tabItems.SelectedItem.Index - 1 <> -1 And tabItems.SelectedItem.Index - 1 <> 10 Then
        frmAdmin.lastIndex = lastTab - 1
        frmAdmin.optCat(tabItems.SelectedItem.Index - 1).Value = True
        frmAdmin.optCat_MouseUp tabItems.SelectedItem.Index - 1, 0, 0, 0, 0
    End If

    Me.Caption = "Item Spawner - " & tabItems.SelectedItem.Caption & " -> " & listItems.listItems.Count & " item" & IIf(listItems.listItems.Count > 1, "s", "") & " available"
    
    lastTab = tabItems.SelectedItem.Index
End Sub

Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 1) As Boolean
    Dim test As textBox, TempValue As String, verified As Byte
    
    If textBox.text = "" Then
        textBox.text = CStr(defaultVal)
        valueToChange = defaultVal
        correctValue = True
    End If

    If Len(textBox.text) = 1 And InStr(1, textBox.text, "-") = 1 Then
        correctValue = True
        cmdSpawn.Enabled = False
        Exit Function
    ElseIf Len(textBox.text) = 1 And IsNumeric(textBox.text) Then
        verified = verifyValue(textBox, min, max)
        If verified = 1 Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        ElseIf verified = 2 Then
            textBox.text = CStr(min)
            textBox.SelStart = 0
            textBox.SelLength = Len(textBox.text)
            correctValue = False
        ElseIf verified = 3 Then
        End If
        cmdSpawn.Enabled = True
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 0 And InStrRev(textBox.text, "-") = 0 And IsNumeric(textBox.text) Then
        verified = verifyValue(textBox, min, max)
        If verified = 1 Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        ElseIf verified = 2 Then
            textBox.text = CStr(min)
            textBox.SelStart = 0
            textBox.SelLength = Len(textBox.text)
            correctValue = False
        ElseIf verified = 3 Then
            textBox.text = CStr(max)
            textBox.SelStart = 0
            textBox.SelLength = Len(textBox.text)
            correctValue = False
        End If
        cmdSpawn.Enabled = True
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 1 And InStrRev(textBox.text, "-") = 1 And IsNumeric(textBox.text) Then
        verified = verifyValue(textBox, min, max)
        If verified = 1 Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        ElseIf verified = 2 Then
            textBox.text = CStr(min)
            textBox.SelStart = 0
            textBox.SelLength = Len(textBox.text)
            correctValue = False
        ElseIf verified = 3 Then
            textBox.text = CStr(max)
            textBox.SelStart = 0
            textBox.SelLength = Len(textBox.text)
            correctValue = False
        End If
        cmdSpawn.Enabled = True
    Else
        textBox.text = CStr(valueToChange)
        textBox.SelStart = Len(textBox.text)
        correctValue = False
    End If
End Function

Private Sub reviseValue(ByRef textBox As textBox, ByRef valueToChange)
    If Not IsNumeric(textBox.text) Then
        textBox.text = CStr(valueToChange)
    Else
        textBox.text = CStr(valueToChange)
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long) As Byte
    Dim Msg As String
    
    If (CDec(txtBox.text) >= min And CDec(txtBox.text) <= max) Then
        verifyValue = 1
    Else
        If CDec(txtBox.text) < min Then
            verifyValue = 2
        Else
            verifyValue = 3
        End If
    End If
End Function

Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub

Private Sub tabItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If FormVisible("frmItemDesc") Then Unload frmItemDesc
End Sub

Private Sub txtAmount_Change()
    If correctValue(txtAmount, currentAmount, 1, currentMaxLimit) Then
        If picked Then
                Me.Caption = "Item Spawner - Going to spawn " & txtAmount.text & " " & listItems.listItems(listItems.SelectedItem.Index).text
        End If
    End If
End Sub

Private Sub txtAmount_Click()
     selectValue txtAmount
End Sub

Private Sub txtAmount_GotFocus()
     selectValue txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    reviseValue txtAmount, currentAmount
End Sub
