EnableExplicit

#WindowNum = 0
#User32dll = 0

#WM_CLIPBOARDUPDATE = $031D

Prototype.l AddClipboardFormatListener(hwnd.i)
Prototype.l OpenClipboard(hWndNewOwner.i)
Prototype.l CloseClipboard()
Prototype.i GetClipboardData(uFormat.l)
Prototype.i SetClipboardData(uFormat.l, hMem.i)
Prototype.l EmptyClipboard()

Global Exit.a = #False
Global ChangedClipboard.a = #False
Global AddClipboardFormatListener_.AddClipboardFormatListener
Global OpenClipboard_.OpenClipboard
Global CloseClipboard_.CloseClipboard
Global GetClipboardData_.GetClipboardData
Global SetClipboardData_.SetClipboardData
Global EmptyClipboard_.EmptyClipboard

Procedure.a IsValidCPF(CPF.s)
  Protected t, d, c
  For t = 9 To 10
    d = 0
    For c = 0 To (t - 1)
      d + Val(Mid(CPF, c + 1, 1)) * ((t + 1) - c)
    Next c
    d = ((10 * d) % 11) % 10;
    If Val(Mid(CPF, c + 1, 1)) <> d
      ProcedureReturn #False
    EndIf
  Next t
  
  ProcedureReturn #True
EndProcedure

Procedure.s ExtractPotentialCPF(PotentialCPF.s)
  PotentialCPF = Trim(PotentialCPF);trim space
  PotentialCPF = Trim(PotentialCPF, Chr($09));trim tab
  PotentialCPF = Trim(PotentialCPF, Chr($0A));trim new line
  PotentialCPF = Trim(PotentialCPF, Chr($0D));trim carriage return
  PotentialCPF = Trim(PotentialCPF, Chr($0B));trim vertical tab
  Protected i
  Protected StringSize = Len(PotentialCPF)
  Protected CPF.s = ""
  For i = 1 To StringSize
    Protected Character.s = Mid(PotentialCPF, i, 1)
    Protected AsciiChar = Asc(Character)
    If AsciiChar > 47 And AsciiChar < 58;filter out non numeric chars
      CPF + Character
    EndIf
    
  Next i
  
  If Len(CPF) = 11 And IsValidCPF(CPF)
    ProcedureReturn CPF
  EndIf
  
  ProcedureReturn ""
  
  
EndProcedure


Procedure.a ProcessClipBoard()
  Protected NumTries.a = 50, i.a
  Protected OpenedClipBoard.l
  For i = 1 To NumTries
    OpenedClipBoard = OpenClipboard_(#Null)
    If OpenedClipBoard <> 0
      Break
    EndIf
    Delay(2)
  Next i
  
  If OpenedClipBoard <> 0
    Protected *TextMemoryHandle = GetClipboardData_(#CF_TEXT)
    PrintN("===============================")
    PrintN(FormatDate("%hh:%ii:%ss", Date()))
    PrintN("called GetClipboardData_:" + Str(*TextMemoryHandle))
    If *TextMemoryHandle = #Null
      PrintN("TextMemory is null, bailed out")
      CloseClipboard_()
      ProcedureReturn #False
    EndIf
    
    Protected *LockedTextMemory = GlobalLock_(*TextMemoryHandle)
    PrintN("called GlobalLock_:" + Str(*LockedTextMemory))
    Protected TextSize.q = GlobalSize_(*TextMemoryHandle)
    PrintN("called globalsize_:" + Str(TextSize))
    If *LockedTextMemory <> #Null
      Protected *Buffer = AllocateMemory(TextSize + 1)
      PrintN("called allocatememory *Buffer is:" + Str(*Buffer))
      ;lstrcpy_(*Buffer, *LockedTextMemory)
      CopyMemory(*LockedTextMemory, *Buffer, TextSize + 1)
      PrintN("called lstrcpy_")
      GlobalUnlock_(*TextMemoryHandle)
      PrintN("called GlobalUnlock_(TextMemory)")
      Protected PotentialCPF.s = PeekS(*Buffer, -1, #PB_Ascii)
      PrintN("called PeekS(*Buffer, -1, #PB_Ascii)")
      FreeMemory(*Buffer)
      PrintN("freed *buffer")
      
      Protected CPF.s = ExtractPotentialCPF(PotentialCPF)
      If Len(CPF) = 0
        CloseClipboard_()
        ProcedureReturn #False
      EndIf
      
      
      
      Protected *CPFBuffer = AllocateMemory(Len(CPF) + 1)
      PrintN("called AllocateMemory(Len(CPF) + 1)")
      PokeS(*CPFBuffer, CPF, 11, #PB_Ascii)
      PrintN("called PokeS(*CPFBuffer, CPF, 11, #PB_Ascii)")
      PrintN("allocated and copied the cpf to *CPFBuffer")
      Protected CPFCopy = GlobalAlloc_(#GMEM_MOVEABLE, Len(CPF) + 1)
      If CPFCopy = #Null
        CloseClipboard_()
        ProcedureReturn #False
      EndIf
      
      PrintN("called globalalloc and cpfcopy is :" + Str(CPFCopy))
      
      Protected CPFCopyLocked = GlobalLock_(CPFCopy)
      PrintN("called GlobalLock_ and CPFCopyLocked is :" + Str(CPFCopyLocked))
      ;lstrcpy_(CPFCopyLocked, *CPFBuffer)
      CopyMemory(*CPFBuffer, CPFCopyLocked, Len(CPF) + 1)
      PrintN("copied from cpfbuffer to cpfcopylocked")
      FreeMemory(*CPFBuffer)
      PrintN("freed *cpfbuffer")
      GlobalUnlock_(CPFCopy)
      PrintN("unlcoked CPFCopy")
      EmptyClipboard_()
      PrintN("calld EmptyClipboard")
      SetClipboardData_(#CF_TEXT, CPFCopy)
      PrintN("called SetClipboardData_")
      CloseClipboard_()
      PrintN("called CloseClipboard_")
      ProcedureReturn #True
    Else
      CloseClipboard_()
      ProcedureReturn #False
    EndIf
    
    CloseClipboard_()
    ProcedureReturn #False
  EndIf
  
  
EndProcedure




If OpenLibrary(#User32dll, "User32.dll") = 0
  MessageRequester("Error!", "Could not open User32.dll!", #PB_MessageRequester_Ok | #PB_MessageRequester_Error)
  End
Else
  AddClipboardFormatListener_ = GetFunction(#User32dll, "AddClipboardFormatListener")
  OpenClipboard_ = GetFunction(#User32dll, "OpenClipboard")
  CloseClipboard_ = GetFunction(#User32dll, "CloseClipboard")
  GetClipboardData_ = GetFunction(#User32dll, "GetClipboardData")
  SetClipboardData_ = GetFunction(#User32dll, "SetClipboardData")
  EmptyClipboard_ = GetFunction(#User32dll, "EmptyClipboard")
EndIf


OpenWindow(#WindowNum, 0, 0, 200, 100, "CPFClip", #PB_Window_SystemMenu | #PB_Window_ScreenCentered)
OpenConsole()

Global ClipboardListenerAdded.l = AddClipboardFormatListener_(WindowID(#WindowNum))


Repeat
  Global Event = WindowEvent()
  Select Event
    Case #PB_Event_CloseWindow
      Exit = #True
    Case #WM_CLIPBOARDUPDATE
      If Not ChangedClipboard
        If ProcessClipBoard()
          ChangedClipboard = #True
        EndIf
      Else
        ChangedClipboard = #False
        
      EndIf
      
      
  EndSelect
  Delay(10)
Until Exit


End