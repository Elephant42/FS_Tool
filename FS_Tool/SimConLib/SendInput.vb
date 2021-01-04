Option Explicit On
Option Strict On

Imports System.CodeDom
Imports System.Runtime.InteropServices

Public Module SendInputModule

#Region "Public Declarations"

    Public Enum KeyStrokeTypeEnum
        Press
        Release
        Full
    End Enum

    Public Enum KeyModifierEnum
        None
        SH
        ALT
        CTL
        SH_CTL
        SH_ALT
        CTL_ALT
        SH_CTL_ALT
    End Enum

    'Scan Codes
    '00 Is normally an error code
    '01 (Esc)
    '02 (1!), 03 (2@), 04 (3#), 05 (4$), 06 (5%E), 07 (6^), 08 (7&), 09 (8*), 0a (9(), 0b (0)), 0c (-_), 0d (=+), 0e (Backspace)
    '0f (Tab), 10 (Q), 11 (W), 12 (E), 13 (R), 14 (T), 15 (Y), 16 (U), 17 (I), 18 (O), 19 (P), 1a ([{), 1b (]})
    '1c (Enter)
    '1d (LCtrl)
    '1e (A), 1f (S), 20 (D), 21 (F), 22 (G), 23 (H), 24 (J), 25 (K), 26 (L), 27 (;), 28 ('")
    '29 (`~)
    '2a (LShift)
    '2b (\|), on a 102-key keyboard
    '2c (Z), 2d (X), 2e (C), 2f (V), 30 (B), 31 (N), 32 (M), 33 (,<), 34 (.>), 35 (/?), 36 (RShift)
    '37 (Keypad-*) Or (*/PrtScn) on a 83/84-key keyboard
    '38 (LAlt), 39 (Space bar),
    '3a (CapsLock)
    '3b (F1), 3c (F2), 3d (F3), 3e (F4), 3f (F5), 40 (F6), 41 (F7), 42 (F8), 43 (F9), 44 (F10)
    '45 (NumLock)
    '46 (ScrollLock)
    '47 (Keypad-7/Home), 48 (Keypad-8/Up), 49 (Keypad-9/PgUp)
    '4a (Keypad--)
    '4b (Keypad-4/Left), 4c (Keypad-5), 4d (Keypad-6/Right), 4e (Keypad-+)
    '4f (Keypad-1/End), 50 (Keypad-2/Down), 51 (Keypad-3/PgDn)
    '52 (Keypad-0/Ins), 53 (Keypad-./Del)
    '54 (Alt-SysRq) on a 84+ key keyboard
    '55 Is less common; occurs e.g. as F11 on a Cherry G80-0777 keyboard, as F12 on a Telerate keyboard, as PF1 on a Focus 9000 keyboard, And as FN on an IBM ThinkPad.
    '56 mostly on non-US keyboards. It Is often an unlabelled key to the left Or to the right of the left Alt key.
    '57 (F11), 58 (F12) both on a 101+ key keyboard
    '59-5a-...-7f are less common. Assignment Is essentially random. Scancodes 55-59 occur as F11-F15 on the Cherry G80-0777 keyboard. Scancodes 59-5c occur on the RC930 keyboard. X calls 5d `KEY_Begin'. Scancodes 61-64 occur on a Telerate keyboard. Scancodes 55, 6d, 6f, 73, 74, 77, 78, 79, 7a, 7b, 7c, 7e occur on the Focus 9000 keyboard. Scancodes 65, 67, 69, 6b occur on a Compaq Armada keyboard. Scancodes 66-68, 73 occur on the Cherry G81-3000 keyboard. Scancodes 70, 73, 79, 7b, 7d occur on a Japanese 86/106 keyboard.
    'e0 1c (Keypad Enter) 		1c (Enter)
    'e0 1d (RCtrl) 		1d (LCtrl)
    'e0 2a (fake LShift) 		2a (LShift)
    'e0 35 (Keypad-/) 		35 (/?)
    'e0 36 (fake RShift) 		36 (RShift)
    'e0 37 (Ctrl-PrtScn) 		37 (*/PrtScn)
    'e0 38 (RAlt) 		38 (LAlt)
    'e0 46 (Ctrl-Break) 		46 (ScrollLock)
    'e0 47 (Grey Home) 		47 (Keypad-7/Home)
    'e0 48 (Grey Up) 		48 (Keypad-8/UpArrow)
    'e0 49 (Grey PgUp) 		49 (Keypad-9/PgUp)
    'e0 4b (Grey Left) 		4b (Keypad-4/Left)
    'e0 4d (Grey Right) 		4d (Keypad-6/Right)
    'e0 4f (Grey End) 		4f (Keypad-1/End)
    'e0 50 (Grey Down) 		50 (Keypad-2/DownArrow)
    'e0 51 (Grey PgDn) 		51 (Keypad-3/PgDn)
    'e0 52 (Grey Insert) 		52 (Keypad-0/Ins)
    'e0 53 (Grey Delete) 		53 (Keypad-./Del) 
    Public Enum ScanCodeEnum
        KEY_NONE = &H0
        KEY_LEFTCTRL = &H1D
        KEY_LEFTSHIFT = &HE2A
        KEY_LEFTALT = &HE38
        KEY_A = &H1F
        KEY_B = &H30
        KEY_C = &H2E
        KEY_D = &H20
        KEY_E = &H12
        KEY_F = &H21
        KEY_G = &H22
        KEY_H = &H23
        KEY_I = &H17
        KEY_J = &H24
        KEY_K = &H25
        KEY_L = &H26
        KEY_M = &H32
        KEY_N = &H31
        KEY_O = &H18
        KEY_P = &H19
        KEY_Q = &H10
        KEY_R = &H13
        KEY_S = &H1F
        KEY_T = &H14
        KEY_U = &H16
        KEY_V = &H2F
        KEY_W = &H11
        KEY_X = &H2D
        KEY_Y = &H15
        KEY_Z = &H2C
        KEY_1 = &H2
        KEY_2 = &H3
        KEY_3 = &H4
        KEY_4 = &H5
        KEY_5 = &H6
        KEY_6 = &H7
        KEY_7 = &H8
        KEY_8 = &H9
        KEY_9 = &HA
        KEY_0 = &HB
        KEY_F1 = &H3B
        KEY_F2 = &H3C
        KEY_F3 = &H3D
        KEY_F4 = &H3E
        KEY_F5 = &H3F
        KEY_F6 = &H40
        KEY_F7 = &H41
        KEY_F8 = &H42
        KEY_F9 = &H43
        KEY_F10 = &H44
        KEY_F11 = &H57
        KEY_F12 = &H58
        KEY_ENTER = &H1C
        KEY_ESC = &H1
        KEY_MINUS = &HC
        KEY_EQUAL = &HD
        KEY_BACKSPACE = &HE
        KEY_TAB = &HF
        KEY_SPACE = &H39
        KEY_GRAVE_TILDE = &H29
        KEY_INSERT = &HE052
        KEY_HOME = &HE047
        KEY_DELETE = &HE053
        KEY_END = &HE04F
        KEY_PAGEUP = &HE049
        KEY_PAGEDOWN = &HE051
        KEY_UP = &HE048
        KEY_RIGHT = &HE04D
        KEY_DOWN = &HE050
        KEY_LEFT = &HE04B
        KEY_NUMLOCK = &H45
        KEY_KPSLASH = &HE035
        KEY_KPASTERISK = &H37
        KEY_KPMINUS = &H4A
        KEY_KPPLUS = &H4E
        KEY_KPENTER = &HE01C
        KEY_KP1 = &H4F
        KEY_KP2 = &H50
        KEY_KP3 = &H51
        KEY_KP4 = &H4B
        KEY_KP5 = &H4C
        KEY_KP6 = &H4D
        KEY_KP7 = &H47
        KEY_KP8 = &H48
        KEY_KP9 = &H49
        KEY_KP0 = &H52
        KEY_KPDEL = &H53
        KEY_LEFTBRACE = &H1A
        KEY_RIGHTBRACE = &H1B
        KEY_BACKSLASH = &H2B
        KEY_SEMICOLON = &H27
        KEY_QUOTE = &H28
        KEY_COMMA = &H33
        KEY_DOT = &H34
        KEY_SLASH = &H35
        KEY_CAPSLOCK = &H3A
    End Enum

    Public Structure KeyEventData
        Public ScanCode As ScanCodeEnum
        Public KeyModifier As KeyModifierEnum
        Public StrokeType As KeyStrokeTypeEnum
    End Structure

#End Region

#Region "Private Declarations"

    Private Const KEYEVENTF_KEYDOWN As Integer = 0
    Private Const KEYEVENTF_KEYUP As Integer = &H2          'If specified, the key Is being released. If Not specified, the key Is being pressed.
    Private Const KEYEVENTF_EXTENDEDKEY As Integer = &H1    'If specified, the scan code was preceded by a prefix Byte that has the value 0xE0 (224).
    Private Const KEYEVENTF_SCANCODE As Integer = &H8       'If specified, wScan identifies the key And wVk Is ignored.
    Private Const KEYEVENTF_UNICODE As Integer = &H4        'If specified, the system synthesizes a VK_PACKET keystroke. The wVk parameter must be zero. This flag can only be combined With the KEYEVENTF_KEYUP flag. For more information, see the Remarks section. 
    Private Const INPUT_MOUSE As Integer = 0
    Private Const INPUT_KEYBOARD As Integer = 1
    Private Const INPUT_HARDWARE As Integer = 2

#If x64 Then
    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Friend Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Friend Structure KEYBDINPUT
        Public wVk As UShort
        Public wScan As UShort
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=8)>
    Friend Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As UShort
        Public wParamH As UShort
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Friend Structure INPUT
        <FieldOffset(0)>
        Public type As Integer
        <FieldOffset(8)>
        Public mi As MOUSEINPUT
        <FieldOffset(8)>
        Public ki As KEYBDINPUT
        <FieldOffset(8)>
        Public hi As HARDWAREINPUT
    End Structure
#Else
    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure

    Private Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Private Structure INPUT
        <FieldOffset(0)>
        Public type As Integer
        <FieldOffset(4)>
        Public mi As MOUSEINPUT
        <FieldOffset(4)>
        Public ki As KEYBDINPUT
        <FieldOffset(4)>
        Public hi As HARDWAREINPUT
    End Structure
#End If

    Private Declare Function SendInput Lib "user32" (ByVal nInputs As Integer, ByVal pInputs() As INPUT, ByVal cbSize As Integer) As Integer
    Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As IntPtr, ByVal idAttachTo As IntPtr, ByVal fAttach As Boolean) As Boolean
    Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As IntPtr, ByVal lpwdProcessId As IntPtr) As IntPtr
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As IntPtr
    Private Declare Auto Function FindWindow Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function GetMessageExtraInfo Lib "user32" () As IntPtr

    Private gInput0(0) As INPUT

#End Region

#Region "Public Functions"

    Public Function GetKeyStrokeTypeByName(ByVal keyStrokeName As String) As KeyStrokeTypeEnum

        Try
            Return CType([Enum].Parse(GetType(KeyStrokeTypeEnum), keyStrokeName), KeyStrokeTypeEnum)
        Catch ex As Exception
            Throw New Exception("Invalid KeyStrokeType: " & keyStrokeName)
        End Try

    End Function

    Public Function GetKeyModifierByName(ByVal keyModName As String) As KeyModifierEnum

        Try
            Return CType([Enum].Parse(GetType(KeyModifierEnum), keyModName), KeyModifierEnum)
        Catch ex As Exception
            Return KeyModifierEnum.None
        End Try

    End Function

    Public Function GetScanCodeByName(ByVal pScanEnumName As String) As ScanCodeEnum

        Try
            Return CType([Enum].Parse(GetType(ScanCodeEnum), pScanEnumName), ScanCodeEnum)
        Catch ex As Exception
            Return ScanCodeEnum.KEY_NONE
        End Try

    End Function

    'Public Sub SendScanCode(ByVal scanCode As Integer)
    '    SendScanCode(scanCode, KeyModifierEnum.None)
    'End Sub
    'Public Sub SendScanCode(ByVal keyEvent As KeyEventStruct)
    '    SendScanCode(keyEvent.ScanCode, keyEvent.KeyModifier)
    'End Sub
    'Public Sub SendScanCode(ByVal scanCode As Integer, ByVal keyModifier As KeyModifierEnum)

    '    SendScanCodeStroke(scanCode, keyModifier, KeyStrokeTypeEnum.Press)
    '    SendScanCodeStroke(scanCode, keyModifier, KeyStrokeTypeEnum.Release)

    '    'Dim mEi As IntPtr = GetMessageExtraInfo()
    '    'Dim dwFlags As Integer = 0

    '    'If keyModifier <> KeyModifierEnum.None Then
    '    '    'press the modifier keys
    '    '    sendModifierKeyPresses(keyModifier, mEi)
    '    'End If

    '    'Dim extendedFlag As Boolean = False
    '    'If (scanCode And &HFF00) = &HE000 Then
    '    '    extendedFlag = True
    '    '    scanCode = CType(scanCode And &HFF, UShort)
    '    'End If

    '    ''press the key
    '    'If extendedFlag Then
    '    '    dwFlags = (KEYEVENTF_KEYDOWN Or KEYEVENTF_SCANCODE Or KEYEVENTF_EXTENDEDKEY)
    '    'Else
    '    '    dwFlags = (KEYEVENTF_KEYDOWN Or KEYEVENTF_SCANCODE)
    '    'End If
    '    'doSendKey(scanCode, dwFlags, mEi)

    '    '' release the key
    '    'If extendedFlag Then
    '    '    dwFlags = (KEYEVENTF_KEYUP Or KEYEVENTF_SCANCODE Or KEYEVENTF_EXTENDEDKEY)
    '    'Else
    '    '    dwFlags = (KEYEVENTF_KEYUP Or KEYEVENTF_SCANCODE)
    '    'End If
    '    'doSendKey(scanCode, dwFlags, mEi)

    '    'If keyModifier <> KeyModifierEnum.None Then
    '    '    'release the modifier keys
    '    '    sendModifierKeyReleases(keyModifier, mEi)
    '    'End If

    'End Sub

    Public Sub SendScanCodeStroke(ByVal scanCode As Integer, ByVal strokeType As KeyStrokeTypeEnum)
        SendScanCodeStroke(scanCode, KeyModifierEnum.None, strokeType)
    End Sub
    Public Sub SendScanCodeStroke(ByVal keyEvent As KeyEventData)
        SendScanCodeStroke(keyEvent.ScanCode, keyEvent.KeyModifier, keyEvent.StrokeType)
    End Sub
    Public Sub SendScanCodeStroke(ByVal scanCode As Integer, ByVal keyModifier As KeyModifierEnum, ByVal strokeType As KeyStrokeTypeEnum)

        If strokeType = KeyStrokeTypeEnum.Full Then
            SendScanCodeStroke(scanCode, keyModifier, KeyStrokeTypeEnum.Press)
            SendScanCodeStroke(scanCode, keyModifier, KeyStrokeTypeEnum.Release)
        Else
            Dim mEi As IntPtr = GetMessageExtraInfo()
            Dim dwFlags As Integer = 0

            Dim extendedFlag As Boolean = False
            If (scanCode And &HFF00) = &HE000 Then
                extendedFlag = True
                scanCode = CType(scanCode And &HFF, UShort)
            End If

            If strokeType = KeyStrokeTypeEnum.Press Then
                'send the modifier keys (for a press they are sent before the main key press)
                sendModifierKeys(keyModifier, mEi, strokeType)

                If extendedFlag Then
                    dwFlags = (KEYEVENTF_KEYDOWN Or KEYEVENTF_SCANCODE Or KEYEVENTF_EXTENDEDKEY)
                Else
                    dwFlags = (KEYEVENTF_KEYDOWN Or KEYEVENTF_SCANCODE)
                End If

                doSendKey(scanCode, dwFlags, mEi)
            ElseIf strokeType = KeyStrokeTypeEnum.Release Then
                If extendedFlag Then
                    dwFlags = (KEYEVENTF_KEYUP Or KEYEVENTF_SCANCODE Or KEYEVENTF_EXTENDEDKEY)
                Else
                    dwFlags = (KEYEVENTF_KEYUP Or KEYEVENTF_SCANCODE)
                End If

                doSendKey(scanCode, dwFlags, mEi)

                'send the modifier keys (for a release they are sent after the main key release)
                sendModifierKeys(keyModifier, mEi, strokeType)
            End If
        End If

    End Sub

    Public Sub SetWindowFocus(ByVal windowText As String)

        Dim hWin As IntPtr = FindWindow(Nothing, windowText)
        SetWindowFocus(hWin)

    End Sub
    Public Sub SetWindowFocus(ByVal windowHandle As IntPtr)

        If Not IntPtr.Zero.Equals(windowHandle) Then
            SetForegroundWindow(windowHandle)
        End If

    End Sub

    Public Sub TestSendInput()

        SetWindowFocus("Untitled - Notepad")

        SendScanCodeStroke(ScanCodeEnum.KEY_H, KeyStrokeTypeEnum.Full)
        SendScanCodeStroke(ScanCodeEnum.KEY_E, KeyStrokeTypeEnum.Full)
        SendScanCodeStroke(ScanCodeEnum.KEY_L, KeyStrokeTypeEnum.Full)
        SendScanCodeStroke(ScanCodeEnum.KEY_L, KeyStrokeTypeEnum.Full)
        SendScanCodeStroke(ScanCodeEnum.KEY_O, KeyStrokeTypeEnum.Full)

    End Sub

#End Region

#Region "Private Functions"

    Private Sub sendModifierKeyPresses(ByVal pMod As KeyModifierEnum, ByVal pEi As IntPtr)
        sendModifierKeys(pMod, pEi, KeyStrokeTypeEnum.Press)
    End Sub

    Private Sub sendModifierKeyReleases(ByVal pMod As KeyModifierEnum, ByVal pEi As IntPtr)
        sendModifierKeys(pMod, pEi, KeyStrokeTypeEnum.Release)
    End Sub

    Private Sub sendModifierKeys(ByVal pMod As KeyModifierEnum, ByVal pEi As IntPtr, ByVal strokeType As KeyStrokeTypeEnum)

        If strokeType = KeyStrokeTypeEnum.Press Then
            sendModifierKeys(pMod, (KEYEVENTF_KEYDOWN Or KEYEVENTF_SCANCODE), pEi)
        Else
            sendModifierKeys(pMod, (KEYEVENTF_KEYUP Or KEYEVENTF_SCANCODE), pEi)
        End If

    End Sub
    Private Sub sendModifierKeys(ByVal pMod As KeyModifierEnum, ByVal keyEventFlags As Integer, ByVal pEi As IntPtr)

        If pMod = KeyModifierEnum.None Then Exit Sub

        If pMod = KeyModifierEnum.SH_CTL_ALT Then
            doSendKey(ScanCodeEnum.KEY_LEFTSHIFT, keyEventFlags, pEi)
            doSendKey(ScanCodeEnum.KEY_LEFTCTRL, keyEventFlags, pEi)
            doSendKey(ScanCodeEnum.KEY_LEFTALT, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.SH_CTL Then
            doSendKey(ScanCodeEnum.KEY_LEFTSHIFT, keyEventFlags, pEi)
            doSendKey(ScanCodeEnum.KEY_LEFTCTRL, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.SH_ALT Then
            doSendKey(ScanCodeEnum.KEY_LEFTSHIFT, keyEventFlags, pEi)
            doSendKey(ScanCodeEnum.KEY_LEFTALT, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.CTL_ALT Then
            doSendKey(ScanCodeEnum.KEY_LEFTCTRL, keyEventFlags, pEi)
            doSendKey(ScanCodeEnum.KEY_LEFTALT, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.SH Then
            doSendKey(ScanCodeEnum.KEY_LEFTSHIFT, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.CTL Then
            doSendKey(ScanCodeEnum.KEY_LEFTCTRL, keyEventFlags, pEi)
        ElseIf pMod = KeyModifierEnum.ALT Then
            doSendKey(ScanCodeEnum.KEY_LEFTALT, keyEventFlags, pEi)
        End If

    End Sub

    Private Sub doSendKey(ByVal scanCode As Integer, ByVal eventFlags As Integer, ByVal pEi As IntPtr)

        loadINPUT(gInput0(0), scanCode, eventFlags, pEi)
        SendInput(1, gInput0, Marshal.SizeOf(GetType(INPUT)))
        doRandomKeySleep()

    End Sub

    Private Sub loadINPUT(ByRef pINP As INPUT, ByVal scanCode As Integer, ByVal pFlags As Integer, ByVal pEi As IntPtr)

        pINP.type = INPUT_KEYBOARD
        pINP.ki.wVk = 0
#If x64 Then
        pINP.ki.wScan = CType(scanCode, UShort)
#Else
        pINP.ki.wScan = CType(scanCode, Short)
#End If
        pINP.ki.dwFlags = pFlags
        'pINP.ki.dwExtraInfo = pEi

    End Sub

    Private Sub doRandomKeySleep()
        Threading.Thread.Sleep(15 + CInt(Rnd() * 15))
        'Threading.Thread.Sleep(50 + CInt(Rnd() * 50))
    End Sub

#End Region

End Module

