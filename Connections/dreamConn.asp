<%
' FileName="default_oledb.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
Dim MM_dreamConn_STRING
MM_dreamConn_STRING = "DSN=dsnBlogMX"

Function RemoveHTML( strText )
    Dim TAGLIST
    TAGLIST = ";!--;!DOCTYPE;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;BDO;BGSOUND;BIG;BODY;BUTTON;CAPTION;CENTER;CITE;COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EMBED;FIELDSET;FONT;FORM;FRAME;FRAMESET;HEAD;H1;H4;H5;H6;HTML;I;IFRAME;INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LINK;LISTING;MAP;MARQUEE;MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OPTGROUP;OPTION;PARAM;PLAINTEXT;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STYLE;SUB;SUP;TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;VAR;WBR;XMP;"

    Const BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"
    
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
	    strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ")

            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
            
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
                bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    strResult = strResult & strText
    
    RemoveHTML = strResult
End Function

%> 