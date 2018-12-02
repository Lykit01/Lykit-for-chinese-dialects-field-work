'2018/08/24
Sub to_ipa_font()
    Dim i As Integer, j As Integer, k As Integer, rowRec As Integer, n As Integer, advanceCol As Integer
    tgSht = 1
    objArr = Sheets(tgSht).Range("a1").CurrentRegion.Value
    For i = 1 To UBound(objArr, 2)
        If objArr(1, i) Like "*记音*" Then advanceCol = i
    Next
    inputStr = InputBox(prompt:="格式：列1，列2...，如5,6,7,8。不输入的话默认为记音的列数", Title:="转写符号转IPA的对象", Default:=advanceCol)
    If inputStr = advanceCol Or inputStr Like "*,*" Then
        inputArr = Split(inputStr, ",")
        For i = 0 To UBound(inputArr)
            col = Val(inputArr(i))
            If col < UBound(objArr, 2) Then
                For j = 2 To UBound(objArr, 1)
                    If objArr(j, col) <> "" Then objArr(j, col) = convert_to_ipa(objArr(j, col))
                Next
            End If
        Next
        Else
            Exit Sub
    End If '方正国际音标默认数字上标，因此不再在程序里将数字上标
    objSht = "ipaF" '要处理的对象
    Sheets(tgSht).Select
    Sheets.Add.Name = objSht
    Sheets(objSht).[a1].Resize(UBound(objArr, 1), UBound(objArr, 2)) = objArr
End Sub

Public Function convert_to_ipa(ByVal str As String)
    Dim c4() As String, c3() As String, c2() As String, c1() As String, v3() As String, v2() As String, v1() As String
    Dim ci4() As String, ci3() As String, ci2() As String, ci1() As String, vi3() As String, vi2() As String, vi1() As String
    Dim vin3() As String, vin2() As String, vin1() As String
    c4 = Split("tsrh,dzrh,tssh,dzzh,tcjh,dzjh", ",")
    ci4 = Split(",,,,,", ",")
    For i = 0 To UBound(c4)
        If InStr(str, c4(i)) > 0 Then str = Replace(str, c4(i), ci4(i))
    Next
    c3 = Split("pfh,bvh,t>h,d>h,tsh,dzh,trh,drh,tsr,dzr,tss,dzz,tjh,djh,tcj,dzj,c!h,h*!", ",")
    ci3 = Split(",,,,,,,,,,,,,,,,,", ",")
    For i = 0 To UBound(c3)
        If InStr(str, c3(i)) > 0 Then str = Replace(str, c3(i), ci3(i))
    Next
    c2 = Split("ph,bh,p*,b*,w$,y$,pf,bv,mg,v$,t>,d>,s>,z>,ts,dz,th,dh,r*,ls,l#,r$,tr,dr,nr,r^,lr,sr,zr,rr,ss,zz,tj,dj,nj,cj,zj,ch,c!,nc,lc,c#,jj,kh,gh,ng,x!,qh,Gh,X!,w!,h*,?h,h!", ",")
    ci2 = Split("p,b,,,w,,pf,,,,,,,,,,t,d,,,,?,,,,,,,,,,,,,,,,c,,,,?,,k,,,,q,,,,,,", ",")
    For i = 0 To UBound(c2)
        If InStr(str, c2(i)) > 0 Then str = Replace(str, c2(i), ci2(i))
    Next
    c1 = Split("p,b,m,f,v,s,z,t,d,n,r,l,c,j,k,g,x,q,G,N,R,X,?,h", ",")
    ci1 = Split("p,b,m,f,v,s,z,t,d,n,,,c,j,k,,x,q,,,,,,h", ",")
    For i = 0 To UBound(c1)
        If InStr(str, c1(i)) > 0 Then str = Replace(str, c1(i), ci1(i))
    Next
    v3 = Split("e+@,e+#,o+$,rowRec>@", ",")
    vi3 = Split(",,,", ",")
    vin3 = Split(",,,", ",")
    For i = 0 To UBound(v3)
        If InStr(str, v3(i) + "~") > 0 Then str = Replace(str, v3(i) + "~", vin3(i))
        If InStr(str, v3(i)) > 0 Then str = Replace(str, v3(i), vi3(i))
    Next
    v2 = Split("i<,i>,y<,y>,i#,u#,u=,e@,e#,o#,e>,e=,e+,o+,rowRec^,A^,rowRec@,rowRec>", ",")
    vi2 = Split(",,,,,,,,,,,,,,,,,", ",")
    vin2 = Split(",,,,,,,,,,,,,,,,,", ",")
    For i = 0 To UBound(v2)
        If InStr(str, v2(i) + "~") > 0 Then str = Replace(str, v2(i) + "~", vin2(i))
        If InStr(str, v2(i)) > 0 Then str = Replace(str, v2(i), vi2(i))
    Next
    v1 = Split("i,y,u,I,Y,U,e,o,E,rowRec,A", ",")
    vi1 = Split("i,y,u,,Y,,e,o,,,A", ",")
    vin1 = Split(",,,,,,,,,,", ",")
    For i = 0 To UBound(v1)
        If InStr(str, v1(i) + "~") > 0 Then str = Replace(str, v1(i) + "~", vin1(i))
        If InStr(str, v1(i)) > 0 Then str = Replace(str, v1(i), vi1(i))
    Next
    convert_to_ipa = str
End Function
