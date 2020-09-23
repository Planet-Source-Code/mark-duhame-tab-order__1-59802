VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSortedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************************
' Name: Alphabetized (sorted) String collection class
' Description:' Instead of creating an Unsorted String
' Array and apply later sorting algorithms on it ,
' we can at once create a sorted collection.
' Every new item who is added in the collection will
' be put in the right position.
' Unlike Arrays the Collections does not
' need Redim Preserve and Shift items.
' Althought it is going to lose the battle in speed
' against a simple array sort,
'
' it can definitely win the war if we put elements in
' the array from time to time and not at once, and we
' need to redim & shift items
'
' By: BAS_ARM
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
' Side Effects:Copy Paste the below in a
'     cls CLASS module
' This code is copyrighted and has limited warranties.
' Please see http://www.Planet-Source-Code.com/xq/ASP/txtCodeId.46418/lngWId.1/qx/vb/scripts/ShowCode.htm
' for details.
'**************************************
'
'---------------------------------------
'---------------------------------------
'
' Module : clsSortedCollection
' DateTime : 24/6/2003 21:00
' Author : Basilis Armeniakos
'---------------------------------------
'----------------------------------------
'
Option Explicit
Private m_Col As Collection
Private m2_Col As Collection
Private k_Col As Collection
Private m_bUniqueValues As Boolean
Private m_iOptionCompare As VbCompareMethod

Private Type MyError
    ErrNumber As Long
    ErrSource As String
    ErrDescription As String
End Type
Private LocalError As MyError
'---------------------------------------
'     ----------------------------------------
'     --------
' Procedure : bUniqueValues
' Arguments :
' Return : Boolean (b)
'---------------------------------------
'     ----------------------------------------
'     --------

Public Property Get bUniqueValues() As Boolean
    
    bUniqueValues = m_bUniqueValues
    
End Property

Public Property Let bUniqueValues(ByVal bbUniqueValues As Boolean)
    
    m_bUniqueValues = bbUniqueValues
    
End Property
'---------------------------------------
'     ----------------------------------------
'     --------
' Procedure : OptionCompare
' Return : (i)
'---------------------------------------
'---------------------------------------
'

Public Property Get OptionCompare() As VbCompareMethod
    
    OptionCompare = m_iOptionCompare
    
End Property

Public Property Let OptionCompare(ByVal x As VbCompareMethod)
    
    m_iOptionCompare = x
    
End Property
'---------------------------------------
'---------------------------------------
'
' Procedure : Class_Initialize
'---------------------------------------
'---------------------------------------
'

Private Sub Class_Initialize()
    
    Set m_Col = New Collection
    Set k_Col = New Collection
    Set m2_Col = New Collection

End Sub
'---------------------------------------
'     ----------------------------------------
'     --------
' Procedure : Class_Terminate
'---------------------------------------
'---------------------------------------
'

Private Sub Class_Terminate()
    
    Set m_Col = Nothing
    Set k_Col = Nothing
    Set m2_Col = Nothing
    
End Sub
'---------------------------------------
'---------------------------------------
'
' Procedure : NewEnum
' Purpose : ' In the Procedure ID box, type –4
' (minus four) to give NewEnum the
' procedure ID required by For Each … Next
' you must also Hide that Member
'
' Return : IUnknown ()
'---------------------------------------
'---------------------------------------
'

Public Property Get NewEnum() As IUnknown
    
    Set NewEnum = m_Col.[_NewEnum]
    
End Property
'---------------------------------------
'---------------------------------------
'
' Procedure : Count
' Return : Long (l)
'---------------------------------------
'---------------------------------------
'

Public Property Get Count() As Long
    
    Count = m_Col.Count
    
End Property
'---------------------------------------
'----------------------------------------
'
' Procedure : Item
' Return : String (s)
'---------------------------------------
'---------------------------------------
'

Public Function Item(swicth As Long, ByVal Lindex As Long, Optional ByVal skey As String) As String
    On Error GoTo ErrorH

    Select Case swicth
        Case 1
            If LenB(skey) Then
                Item = m_Col.Item(skey)
            Else ' ByIndex
                Item = m_Col.Item(Lindex)
            End If
        Case 2
            If LenB(skey) Then
                Item = k_Col.Item(skey)
            Else ' ByIndex
                Item = k_Col.Item(Lindex)
            End If
        Case 3
            If LenB(skey) Then
                Item = m2_Col.Item(skey)
            Else
                Item = m2_Col.Item(Lindex)
            End If
    End Select
    
    Exit Function
ErrorH:
    LocalError = ErrorRaiser(Lindex, skey)

    With LocalError
        Err.Raise .ErrNumber, .ErrSource, .ErrDescription
    End With
    
End Function
'---------------------------------------
'---------------------------------------
'
' Procedure : Remove
' Purpose : Removes the specified by index, or by
' key item Invalid index or key RaiseError
' Arguments : (Optional ByVal Lindex As
'     Long, Optional ByVal skey As String)
' Return : String (s)
'---------------------------------------
'---------------------------------------
'
' For Removing A Specific Value try the
' L = ItemExists(sItem)
' Do While L
' Remove(L)
' Loop

Public Sub Remove(Optional ByVal Lindex As Long, Optional ByVal skey As String)
    On Error GoTo ErrorH

    If LenB(skey) Then
        m_Col.Remove skey
        m2_Col.Remove skey
        k_Col.Remove Lindex
    Else ' ByIndex
        m_Col.Remove Lindex
        m2_Col.Remove Lindex
        k_Col.Remove Lindex
    End If
    
    Exit Sub
ErrorH:
    LocalError = ErrorRaiser(Lindex, skey)

    With LocalError
        Err.Raise .ErrNumber, .ErrSource, .ErrDescription
    End With
    
End Sub
'---------------------------------------
'---------------------------------------
'
' Procedure : Add
' Purpose : Stores the item in the proper position
'
' Arguments : (ByVal Sitem As String)
' Return : ()
'
'---------------------------------------
'---------------------------------------

Public Sub Add(ByVal S As String, ByVal S2 As String, Optional ByVal skey As String)
    Dim L As Long
    Dim U As Long
    Dim M As Long
    Dim icomp As Long
    Dim Sitem As String
    Dim sitem2 As String
    
    If m_Col.Count = 0 Then
        m_Col.Add S
        m2_Col.Add S2
        k_Col.Add skey
    Else
        L = 1
        U = m_Col.Count

        If U = 1 Then
            M = U ' = 1
            Sitem = m_Col.Item(M)
            sitem2 = m2_Col.Item(M)
            If Val(Sitem) = Val(S) Then
                If Val(sitem2) = Val(S2) Then
                    icomp = 0
                ElseIf Val(sitem2) < Val(S2) Then
                    icomp = -1
                ElseIf Val(sitem2) > Val(S2) Then
                    icomp = 1
                End If
            ElseIf Val(Sitem) < Val(S) Then
                icomp = -1
            ElseIf Val(Sitem) > Val(S) Then
                icomp = 1
            End If
        Else
            Do While U >= L
                M = (L + U) \ 2
                Sitem = m_Col.Item(M)
                sitem2 = m2_Col.Item(M)
                If Val(Sitem) = Val(S) Then
                    If Val(sitem2) = Val(S2) Then
                        icomp = 0
                    ElseIf Val(sitem2) < Val(S2) Then
                        icomp = -1
                    ElseIf Val(sitem2) > Val(S2) Then
                        icomp = 1
                    End If
                ElseIf Val(Sitem) < Val(S) Then
                    'And Val(sitem2) < Val(S2) Then
                    icomp = -1
                ElseIf Val(Sitem) > Val(S) Then
                    'And Val(sitem2) > Val(S2) Then
                    icomp = 1
                End If

                Select Case icomp
                    Case -1
                    L = M + 1
                    Case 0

                    If m_bUniqueValues Then ' if we want only unique value to be stored in the collection
                        Exit Sub
                    End If
                    Exit Do
                    
                    Case 1 ' sItem > s so we must search in the low half
                    U = M - 1
                End Select
            Loop
        End If
        If LenB(skey) Then
            If icomp = 1 Then ' before
                m_Col.Add S, skey, before:=M
                m2_Col.Add S2, skey, before:=M
                k_Col.Add skey, skey, before:=M
            Else ' after M
                m_Col.Add S, skey, after:=M
                m2_Col.Add S2, skey, after:=M
                k_Col.Add skey, skey, after:=M
            End If 'If icomp = 1
        Else 'If LenB(skey) skey is missing
            If icomp = 1 Then ' before
                m_Col.Add S, before:=M
                m2_Col.Add S2, before:=M
                k_Col.Add skey, before:=M
            Else ' after M
                m_Col.Add S, after:=M
                m2_Col.Add S2, after:=M
                k_Col.Add skey, after:=M
            End If 'If icomp = 1
        End If 'If LenB(skey)
    End If 'If m_Col.Count = 0 Then

End Sub
'---------------------------------------
'---------------------------------------
'
' Procedure : ExistsInCollection
' Purpose : returns the index if it is found
' in the collection, or 0 if not
'---------------------------------------
'---------------------------------------
'
' iMode = 0 => V = sItem
' iMode = 1 => V = skey
' iMode = 2 => V = Index

Public Function ExistsInCollection(ByVal S As String, Optional ByVal iMode As Long) As Long
    Dim L As Long
    Dim U As Long
    Dim M As Long
    Dim icomp As Long
    Dim Sitem As String
    
    On Error GoTo ErrorH

    If m_Col.Count Then
        Select Case iMode
            Case 0
                L = 1
                U = m_Col.Count

                Do While U >= L ' again the same, but to put that in a separate procedure
                    M = (L + U) \ 2
                    Sitem = m_Col.Item(M)
                    icomp = StrComp(Sitem, S, m_iOptionCompare)

                    Select Case icomp
                        Case -1 ' sItem < s
                        L = M + 1
                        Case 0
                        ExistsInCollection = M
                        Exit Function
                        Case 1 '' sItem > s so we must search in the low half
                        U = M - 1
                    End Select
                Loop
            Case 1
                L = CLng(S) ' that it can be done also with many other ways
                m_Col.Item L 'if L is invalid an error occures
                ExistsInCollection = L
                Case 2
                Sitem = m_Col.Item(S) 'If key does not exists an error occures
                L = ExistsInCollection(Sitem, 0) ' the second time it will run for searching the stored item, so 2 runs with one DoLoop the second time _
                instead of having separate procedures
                ExistsInCollection = L
        End Select
    End If
ErrorH:
'Simply Returns 0

End Function
'---------------------------------------
'---------------------------------------
'
' Procedure : ErrorRaiser
' Purpose : Testing for Errors in the ar
'     guments Line
' Return : MyError ()
'---------------------------------------
'---------------------------------------
'

Private Function ErrorRaiser(Optional ByVal Lindex As Long, Optional ByVal skey As String) As MyError
    Dim LerrNum As Long
    Dim SerrDesc As String

    If m_Col.Count = 0 Then
        LerrNum = 6000
        SerrDesc = "Empty Collection"
    Else 'If m_Col.Count = 0
        If LenB(skey) Then
            On Error Resume Next
            m_Col.Item skey 'test if the inserted key is valid
            If Err.Number Then
                LerrNum = 6001
                SerrDesc = "Not Matching Key"
                Err.Clear
                On Error GoTo 0
            End If 'If Err.Number
        Else 'If LenB(skey)
            If Lindex = 0 Then
                LerrNum = 6002
                SerrDesc = "Key And Index are both Missing"
            Else 'If Lindex = 0
                On Error Resume Next
                m_Col.Item Lindex 'test if the inserted key is valid
                If Err.Number Then
                    LerrNum = 6003
                    SerrDesc = "Invalid Index"
                    Err.Clear
                    On Error GoTo 0
                End If 'If Err.Number
            End If 'If Lindex = 0
        End If 'If LenB(skey)
    End If ' If m_Col.Count = 0

    With ErrorRaiser
        .ErrNumber = vbObjectError + LerrNum
        .ErrDescription = SerrDesc
        .ErrSource = TypeName(Me)
    End With
    
End Function
