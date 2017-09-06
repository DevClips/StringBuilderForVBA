'************************************************************************
'*-----------------------------------------------------------------------
'*  Name： StringBuilder (Class Module)
'*-----------------------------------------------------------------------
'*  Descriptioin：StringBuilder for VBA
'*-----------------------------------------------------------------------
'*  Copyright: HAYs  http://dev-clips.com , 2015 All Rights Reserved.
'*-----------------------------------------------------------------------
'*  <Update>
'*  Date        Version     Author     Memo
'*-----------------------------------------------------------------------
'*  2015.11.25  1.00        HAYs       New Release
'************************************************************************
' option
Option Explicit

'************************************************************************
'*  variable
'************************************************************************
Private pCapacity As Long
Private pLength As Long

Private mBuffer As String

'************************************************************************
'*  class event
'************************************************************************
'*-----------------------------------------------------------------------
'*  constructor
'*-----------------------------------------------------------------------
Private Sub Class_Initialize()
    pCapacity = 1023
    Me.Clear
End Sub

'*-----------------------------------------------------------------------
'*  destructor
'*-----------------------------------------------------------------------
Private Sub Class_Terminate()
    'clean up
    mBuffer = vbNullString
End Sub

'************************************************************************
'*  property
'************************************************************************
'*-----------------------------------------------------------------------
'*  Capacity
'*-----------------------------------------------------------------------
Friend Property Let Capacity(ByVal NewValue As Long)
    'ignore smaller NewValue
    If NewValue > pCapacity Then
        're-allocate
        mBuffer = mBuffer & String(NewValue - pCapacity, vbNullChar)
        'save new value
        pCapacity = NewValue
    End If
End Property
Friend Property Get Capacity() As Long
    Capacity = pCapacity
End Property

'*-----------------------------------------------------------------------
'*  Length
'*-----------------------------------------------------------------------
Friend Property Let Length(ByVal NewValue As Long)
    If NewValue < pLength Then
        Mid(mBuffer, NewValue + 1, pLength - NewValue) = _
            String$(pLength - NewValue, vbNullChar)
    End If
    pLength = NewValue
End Property
Friend Property Get Length() As Long
    Length = pLength
End Property


'************************************************************************
'*  method
'************************************************************************
'*-----------------------------------------------------------------------
'*  clear
'*-----------------------------------------------------------------------
Friend Function Clear() As StringBuilder
    'initialize length
    pLength = 0
    'allocate memory
    mBuffer = String$(pCapacity, vbNullChar)
    'return me
    Set Clear = Me
End Function

'*-----------------------------------------------------------------------
'*  append
'*-----------------------------------------------------------------------
Friend Function Append(ByRef StringValue As String) As StringBuilder
    Dim pos As Long
    Dim tmpCap As Long
    'set position
    pos = pLength + 1
    'add new length
    pLength = pLength + Len(StringValue)
    'check overflow
    If pLength > pCapacity Then
        'expand capacity *doubles up
        tmpCap = pCapacity
        Do While tmpCap < pLength
            tmpCap = tmpCap * 2
        Loop
        'save new capacity
        Me.Capacity = tmpCap
    End If
    'append
    Mid(mBuffer, pos) = StringValue
    'retrun me
    Set Append = Me
End Function

'*-----------------------------------------------------------------------
'*  insert
'*-----------------------------------------------------------------------
Friend Function Insert(ByRef StringValue As String, _
                        ByVal position As Long) As StringBuilder
    Dim tmpCap As Long
    Dim tmpLen As Long
    'check position
    Select Case position
        Case 1 To pLength
        Case Is < 1: position = 1
        Case Else
            Set Insert = Append(StringValue)
            Exit Function
    End Select
    'save length
    tmpLen = pLength
    'add new length
    pLength = pLength + Len(StringValue)
    'check overflow
    If pLength > pCapacity Then
        'expand Capacity *doubles up
        tmpCap = pCapacity
        Do While tmpCap < pLength
            tmpCap = tmpCap * 2
        Loop
        'save new capacity
        Me.Capacity = tmpCap
    End If
    'slide
    Mid(mBuffer, position + Len(StringValue)) _
        = Mid$(mBuffer, position, tmpLen)
    'insert
    Mid(mBuffer, position) = StringValue
    'retrun me
    Set Insert = Me
End Function

'*-----------------------------------------------------------------------
'*  string value
'*-----------------------------------------------------------------------
Friend Function ToString() As String
    ToString = Left$(mBuffer, pLength)
End Function
