Enumeration
  #MAIN_WINDOW
EndEnumeration

Enumeration
  #MENU_BAR
EndEnumeration

Enumeration
  #ACTION_EXIT
  #ACTION_ABOUT
EndEnumeration

Enumeration
  #ZIP_UNPACK_BUTTON
  #ZIP_PACK_BUTTON
  #RAR_UNPACK_BUTTON
  #PPTX_UNPACK_BUTTON
EndEnumeration

Procedure OpenMainWindow()
  
  If OpenWindow(#MAIN_WINDOW, #PB_Ignore, #PB_Ignore, 270, 220, "Архиватор", #PB_Window_MinimizeGadget | #PB_Window_ScreenCentered)
    If CreateMenu(#MENU_BAR, WindowID(#MAIN_WINDOW)) 
      MenuTitle("&Файл")
      MenuItem(#ACTION_EXIT, "В&ыход")
      
      MenuTitle("&Справка")
      MenuItem(#ACTION_ABOUT, "&О программе")
    EndIf
    
    ButtonGadget(#ZIP_UNPACK_BUTTON, 70, 20, 130, 25, "Распаковать ZIP")
    ButtonGadget(#ZIP_PACK_BUTTON, 70, 65, 130, 25, "Запаковать ZIP")
    ButtonGadget(#RAR_UNPACK_BUTTON, 70, 110, 130, 25, "Распаковать RAR")
    ButtonGadget(#PPTX_UNPACK_BUTTON, 70, 155, 130, 25, "Распаковать PPTX")
    
    SetWindowColor(#MAIN_WINDOW, $FFFFFF)
    
    SetGadgetColor(#ZIP_UNPACK_BUTTON, #PB_Gadget_BackColor, $FFFFFF)
    SetGadgetColor(#ZIP_PACK_BUTTON, #PB_Gadget_BackColor, $FFFFFF)
    SetGadgetColor(#RAR_UNPACK_BUTTON, #PB_Gadget_BackColor, $FFFFFF)
    SetGadgetColor(#PPTX_UNPACK_BUTTON, #PB_Gadget_BackColor, $FFFFFF)
  EndIf
EndProcedure

Procedure UnpackZIP()
  archiveFilePath.s = OpenFileRequester("", "", "Архив ZIP (*.zip)|*.zip", 0)
  If archiveFilePath <> ""
    folder.s = ReplaceString(GetFilePart(archiveFilePath), ".zip", "")
    If PureZIP_ExtractFiles(archiveFilePath, "*.*", GetPathPart(archiveFilePath) + folder, #True)
      MessageRequester("Готово", "Архив распакован", #MB_ICONINFORMATION)
    Else
      MessageRequester("Готово", "Архив не распакован", #MB_ICONINFORMATION)
    EndIf
  EndIf
EndProcedure

Procedure PackZIP()
  filePath.s = OpenFileRequester("", "", "Все файлы (*.*)|*.*", 0, #PB_Requester_MultiSelection)
  If filePath <> ""
    If PureZIP_Archive_Create(GetPathPart(filePath) + GetFilePart(filePath) + ".zip", #APPEND_STATUS_CREATE)
      While filePath
        If PureZIP_Archive_Compress(filePath, #False) <> #Z_OK
          MessageRequester("", "Ошибка при добавлении: " + filePath, #MB_ICONINFORMATION)
        EndIf
        filePath = NextSelectedFileName()
      Wend
      PureZIP_Archive_Close()
      MessageRequester("Готово", "Архив создан", #MB_ICONINFORMATION)
    Else
      MessageRequester("Готово", "Архив не создан", #MB_ICONINFORMATION)
    EndIf
  EndIf
EndProcedure

Procedure UnpackRAR()
  archiveFilePath.s = OpenFileRequester("", "", "WinRAR (*.rar)|*.rar", 0)
  If archiveFilePath <> ""
    hRAR = RAR_OpenArchive(archiveFilePath)
    If hRAR
      buffer.s {270}
      While RAR_ReadHeader(hRAR, @buffer) = 0
        RAR_SetPassword(hRAR, "")
        folder.s = ReplaceString(GetFilePart(archiveFilePath), ".rar", "")
        RAR_ProcessFile(hRAR, #RAR_EXTRACT, "", GetPathPart(archiveFilePath) + folder + "\" + buffer)
      Wend
      RAR_CloseArchive(hRAR)
      MessageRequester("Готово", "Архив распакован", #MB_ICONINFORMATION)
    EndIf
  EndIf
EndProcedure

Procedure UnpackPPTX()
  archiveFilePath.s = OpenFileRequester("", "", "Презентация PPTX (*.pptx)|*.pptx", 0)
  If archiveFilePath <> ""
    folder.s = ReplaceString(GetFilePart(archiveFilePath), ".pptx", "")
    If PureZIP_ExtractFiles(archiveFilePath, "*.*", GetPathPart(archiveFilePath) + folder, #True)
      MessageRequester("Готово", "Презентация распакована", #MB_ICONINFORMATION)
    Else
      MessageRequester("Готово", "Презентация не распакована", #MB_ICONINFORMATION)
    EndIf
  EndIf
EndProcedure

OpenMainWindow()

Repeat
  event       = WaitWindowEvent()
  eventMenu   = EventMenu()
  eventGadget = EventGadget()
  eventWindow = EventWindow()
  eventType   = EventType()
  
  If eventWindow = #MAIN_WINDOW
    If event = #PB_Event_Menu
      If eventMenu = #ACTION_EXIT
        Break
      ElseIf eventMenu = #ACTION_ABOUT
        MessageRequester("О программе", "Архиватор. Версия 1.0" + #CR$ + #CR$ + "Автор: Салават Даутов" + #CR$ + #CR$ + "Дата создания: июль 2012", #MB_ICONINFORMATION)
      EndIf
    EndIf
    
    If event = #PB_Event_Gadget
      If eventGadget = #ZIP_UNPACK_BUTTON
        UnpackZIP()
      ElseIf eventGadget = #ZIP_PACK_BUTTON
        PackZIP()
      ElseIf eventGadget = #RAR_UNPACK_BUTTON
        UnpackRAR()
      ElseIf eventGadget = #PPTX_UNPACK_BUTTON
        UnpackPPTX()
      EndIf
    EndIf
  EndIf
Until event = #PB_Event_CloseWindow And eventWindow = #MAIN_WINDOW
