      *%METADATA                                                                          *
      * %TEXT Functions XML Spreadshett                                                   *
      *%EMETADATA                                                                         *


     h nomain   decedit('.')

      /COPY ../QPROTOSRC/XSS.RPGLEINC
      /COPY ../QPROTOSRC/TXT.RPGLEINC

     D init            Pr
     D addEntry        Pr
     d Addstyle        pr
     d Strstyle        pr
     D SetAutoFilter   pr
     D Addcell         pr
     D style                          3u 0 VALUE  OPTIONS(*NOPASS)
     D idx                            5u 0 VALUE  OPTIONS(*NOPASS)
     D merge                          3u 0 VALUE  OPTIONS(*NOPASS)

     D replacement     ds
     D  aa                           10    dim(64) ctdata perrcd(1)
     d   char                         1    overlay(aa)
     d   unicode                      4    overlay(aa:*next)
     D chgvar          pr          2000
     D  var                           1    value
     D  input                         4    value
     D  string                     1024    value
     D chgvar2         pr          2000
     D  var                           1    value
     D  input                         5    value
     D  string                     1024    value

     d style           ds                  qualified inz
     d  color                         7
     d  forcolor                      7
     d  align                         8
     d  valign                        8
     d  font                         20
     d  family                       10
     d  bold                           n
     d  wrap                           n
     d  italic                         n
     d  size                          2  0
     d  format                       20
     d  Bord_B                        2  0
     d  Bord_T                        2  0
     d  Bord_R                        2  0
     d  Bord_L                        2  0

     d user            s             10    inz(*user)
     d OpenRow         s               n
     d OpenStyle       s               n
     d OpenSheet       s               n
     d Autofilter      s               n
     d FreezePane      s               n
     d idx             s              3u 0
     d colId           s              5u 0
     d RowId           s              5u 0
     d Range1          s             12
     d Range2          s             12
     d FreezeH         s              5
     d FreezeV         s              5
     d data            s           2000    varying
     d data2           s           2000
     d EOR             c                   X'0D25'
     d Hash            c                   %ucs2('£')
     d HashUcs         s              1c
     d HashCar         s              1

     PXssopenfile      b                   export
     Dxssopenfile      Pi
     D  File                       1024A   value
      /free
        init();
        txtcrtfile(file:*off);
        data =  '<?xml version="1.0"?>'  + EOR +
            '<?mso-application progid="Excel.Sheet"?>'  + EOR +
            '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"'
              + EOR +
            'xmlns:o="urn:schemas-microsoft-com:office:office"'  + EOR +
            'xmlns:x="urn:schemas-microsoft-com:office:excel"'  + EOR +
            'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' + EOR +
            'xmlns:html="http://www.w3.org/TR/REC-html40">' ;
         addEntry();
      /end-free
     PXssopenfile      e
     p XssSetDocProperties...
     p                 b                   export
     D XssSetDocProperties...
     D                 Pi
     D  author                       20    value
     D  LastAuthor                   20    value
     D  version                       4  2 value
      /free
         data = '<DocumentProperties xmlns="urn:' +
                'schemas-microsoft-com:office:office">';
         addEntry();
         if author = ' ';
           data =  '<Author>' + user  + '</Author>';
          else;
           data =  '<Author>' + %trim(author)  + '</Author>';
          endif;
         addEntry();
         if Lastauthor <> ' ';
           data =  '<LastAuthor>' + %trim(lastAuthor) + '</LastAuthor>';
           addEntry();
         endif;
         data = '<Created>' + %char(%date():*iso) + 'T' +
                 %char(%time():*hms) + 'Z</Created>';
         addEntry();
         if version > 0;
           data = '<Version>' + %char(version) + '</Version>' ;
           addEntry();
         endif;
         data = '</DocumentProperties>';
         addEntry();
      /end-free
     p                 e
     pxsscreatestyle   b                   export
     Dxsscreatestyle   Pi             3u 0
      /free
        if openStyle = *off;
          openStyle = *on;
          data = '<Styles>';
          addEntry();
        else;
          Addstyle();
        ENDIF;
        idx += 1;
        clear style;
        return idx;
      /end-free
     pxsscreatestyle   e

     p Addstyle        b
     d Addstyle        pi
      /free
        if idx > 0;
          data = ' <Style ss:ID="s' + %char(idx) + '">' ;
        else;
          data = ' <Style ss:ID="Default" ss:Name="Normal">' ;
        endif;
        addEntry();
        if style.align <> ' '  or  style.Valign <> ' '
                               or style.Wrap ;
          data = '<Alignment ';
          if  style.align <> ' ';
             data += 'ss:Horizontal="'
                  + %trim(style.align) + '" ';
          ENDIF;
          if  style.Valign <> ' ';
             data += 'ss:Vertical="'
                  + %trim(style.Valign) + '" ';
          ENDIF;
          if style.Wrap;
            data += 'ss:WrapText="1" ';
          ENDIF;
          data += '/>';
          addEntry();
        ENDIF;

        if style.font <> ' ' or  style.family <> ' '
           or  style.color <> ' '  or style.Italic
           or style.size <> 0 or style.Bold;
          data = '<Font ';
          if style.font <> ' ';
            data += 'ss:FontName="' + %trim(style.font) + '" ';
          ENDIF;
           if style.family <> ' ';
            data += 'x:Family="' + %trim(style.family) + '" ';
          ENDIF;
          if style.color <> ' ';
            data += 'ss:Color="' + style.color + '" ';
          ENDIF;
          if style.size <> 0;
            data += 'ss:Size="' + %char(style.size) + '" ';
          ENDIF;
          if style.bold;
            data += 'ss:Bold="1" ';
          ENDIF;
          if style.Italic;
            data += 'ss:Italic="1" ';
          ENDIF;
          data += '/>';
          addEntry();
        ENDIF;
        if style.forcolor <> ' ';
          data = '<Interior ss:Color="' +  style.forcolor +
                  '" ss:Pattern="Solid"/>';
          addEntry();
        ENDIF;
        if style.format <> ' ';
          data = '<NumberFormat ss:Format="' + %trim(style.format) + '"/>';
          addEntry();
        endif;
        if Style.Bord_B > 0 or Style.Bord_T > 0  or
           Style.Bord_R > 0 or Style.Bord_L > 0 ;
           data = '<Borders>';
           addEntry();
           if Style.Bord_B > 0;
             data = '<Border ss:Position="Bottom" ss:LineStyle="Continuous" '
                     + 'ss:Weight="' + %char(Style.Bord_B) + '"/>';
             addEntry();
           endif;
           if Style.Bord_L > 0;
             data = '<Border ss:Position="Left" ss:LineStyle="Continuous" '
                     + 'ss:Weight="' + %char(Style.Bord_L) + '"/>';
             addEntry();
           endif;
           if Style.Bord_R > 0;
             data = '<Border ss:Position="Right" ss:LineStyle="Continuous" '
                     + 'ss:Weight="' + %char(Style.Bord_R) + '"/>';
             addEntry();
           endif;
           if Style.Bord_T > 0;
             data = '<Border ss:Position="Top" ss:LineStyle="Continuous" '
                     + 'ss:Weight="' + %char(Style.Bord_T) + '"/>';
             addEntry();
           endif;
           data = '</Borders>';
           addEntry();
        ENDIF;
        data = ' </Style>';
        addEntry();
      /end-free
     p                 e
     p strstyle        b
     d Strstyle        pi
      /free
        if idx = 0 and not openStyle;
          data = ' <Styles>' ;
          addEntry();
          openStyle = *on;
        endif;
      /end-free
     p                 e

     pXsssetbold       b                   export
     DXsssetbold       Pi
      /free
        StrStyle();
        style.bold = *on;
      /end-free

     pXsssetbold       e

     pXssSetItalic     b                   export
     DXssSetItalic     Pi

      /free
         StrStyle();
         style.italic = *on;
      /end-free
     p                 e

     pxssSetFontSize   b                   export
     dxssSetFontSize   pi
     d  size                          5i 0 value

      /free
         StrStyle();
         Style.size = size;
      /end-free
     p                 e
     pXssSetFontName   b                   export
     DXssSetFontName   Pi
     d  Name                         50    value
     d  family                       50    value options(*nopass)

      /free
         StrStyle();
         style.font = name;
         if %parms > 1;
            style.family = family;
         ENDIF;
      /end-free
     p                 e
     pXssSetFormat     b                   export
     DXssSetFormat     Pi
     d  format                       20    value

      /free

         StrStyle();
         Format = %xlate('q':Hashcar:format);
         style.format = format;
      /end-free
     p                 e
     pXssSetalign      b                   export
     DXssSetalign      Pi
     d  alignment                    10    value

      /free
         StrStyle();
         style.align = alignment;
      /end-free
     p                 e
     pXsssetVertAlign  b                   export
     dXsssetVertAlign  pi
     d  alignment                    10    value

      /free
         StrStyle();
         style.Valign = alignment;
      /end-free
     p                 e
     pXsssetcolor      b                   export
     DXsssetcolor      Pi
     d  color                         6    value

      /free
         StrStyle();
         style.color = HashCar + color;
      /end-free
     p                 e
     pXsssetForcolor   b                   export
     DXsssetForcolor   Pi
     d  color                         6    value

      /free
         StrStyle();
         style.Forcolor = HashCar + color;
      /end-free
     p                 e
     pXsssetwrap       b                   export
     DXsssetwrap       Pi

      /free
         StrStyle();
         Style.wrap = *on;
      /end-free
     p                 e
     pXsssetBorder     b                   export
     DXsssetBorder     Pi
     d  Weight                        2  0 value
     d  Position                      1    value OPTIONS(*NOPASS)
      /free
         StrStyle();
         if %parms = 1;
           Style.Bord_B = Weight;
           Style.Bord_R = Weight;
           Style.Bord_L = Weight;
           Style.Bord_T = Weight;
           return;
         ENDIF;
         if %scan('B':Position) > 0 ;
           Style.Bord_B = Weight;
         ENDIF;
         if %scan('R':Position) > 0 ;
           Style.Bord_R = Weight;
         ENDIF;
         if %scan('L':Position) > 0 ;
           Style.Bord_L = Weight;
         ENDIF;
         if %scan('T':Position) > 0 ;
           Style.Bord_T = Weight;
         ENDIF;
      /end-free
     p                 e

     pXssaddsheet      b                   export
     DXssaddsheet      Pi
     D sheet                        100A   value OPTIONS(*NOPASS)
      /free
         if OpenStyle;
           Addstyle();
           data = '</Styles>';
           AddEntry();
           openStyle = *off;
         ENDIF;
         If OpenRow;
           data = '</Row>';
           addEntry();
         ENDIF;
         if OpenSheet;
           data = '</Table>';
           AddEntry();
           if Autofilter;
             SetAutoFilter();
           ENDIF;
           if FreezePane;
             SetFreezePane();
           ENDIF;
           data = '</Worksheet>';
           AddEntry();
         ENDIF;
         if %parms = 0 or sheet = ' ';
           data = ' <Worksheet ss:Name="Sheet1">';
         else;
           data = ' <Worksheet ss:Name="' + %trim(sheet) + '">';
         endif;
         AddEntry();
         OpenSheet = *on;
         data = '<Table>';
         AddEntry();
         openrow = *off;
         RowId = 0;
         AutoFilter = *off;
         clear range1;
         clear range2;
         FreezePane = *off;
         clear FreezeH;
         clear FreezeV;
      /end-free
     p                 e
     PxssSetColWidth   B                   Export
     DxssSetColWidth   Pi
     D Width                          5  2 Const
     D ColId                          5i 0 CONST OPTIONS(*NOPASS)
      /free
         If %parms = 2;
           Data = '<Column ss:Index="' + %char(ColId)
                + '" ss:Width="' + %char(Width) + '"/>';
         else;
           Data = '<Column ss:Width="' + %char(Width) + '"/>';
         ENDIF;
         AddEntry();
      /end-free
     p                 e

     pxssaddrow        b                   export
     Dxssaddrow        Pi
     D height                         5  2 CONST OPTIONS(*NOPASS)
      /free
         If OpenRow;
           data = '</Row>';
           addEntry();
         ENDIF;
         if %parms = 0 or Height = 0;
           data = '<Row>';
         else;
          data = '<Row ss:AutoFitHeight="0" ss:Height="' + %char(height) + '">';
         ENDIF;
         addEntry();
         OpenRow = *on;
         RowId += 1;
         ColId = 0;
      /end-free
     p                 e
     p AddCell         b
     D AddCell         Pi
     D style                          3u 0 VALUE  OPTIONS(*NOPASS)
     D idx                            5u 0 VALUE  OPTIONS(*NOPASS)
     D merge                          3u 0 VALUE  OPTIONS(*NOPASS)
      /free
        if %parms = 0;
          data = '<Cell>';
          ColId += 1;
        else;
          data = '<Cell ';
          if %parms >= 1;
            data += ' ss:StyleID="s' + %char(style) + '" ';
          ENDIF;
          if %parms >= 2 and idx > 0;
            data += ' ss:Index="' + %char(idx) + '" ';
            colId = idx;
          else;
            colId += 1;
          ENDIF;
          if %parms = 3 and merge > 0;
            data += ' ss:MergeAcross="' + %char(merge) + '" ';
            colId += merge;
          ENDIF;
          data += '>';
        ENDIF;

      /end-free
     p                 e
     pxssaddcellChar   b                   export
     DxssaddcellChar   Pi
     D value                       1024A   value
     D style                          3u 0 VALUE  OPTIONS(*NOPASS)
     D idx                            5u 0 VALUE  OPTIONS(*NOPASS)
     D merge                          3u 0 VALUE  OPTIONS(*NOPASS)

     D i               s              3u 0
      /free
        if %parms = 1;
          Addcell();
        elseif %parms = 2;
          Addcell(style);
        elseif %parms = 3;
          Addcell(style:idx);
        elseif %parms = 4;
          Addcell(style:idx:merge);
        ENDIF;
        value = chgvar2('&':'&amp;':value);
        value = chgvar2('<':'&lt;':value);
        value = chgvar2('>':'&gt;':value);
        for i = 1 to %elem(aa);
          value = chgvar(char(i):unicode(i):value);
        endfor;
        data += '<Data ss:Type="String">' + %trim(value) +
               	 '</Data></Cell>';
        addEntry();
      /end-free
     p                 e
     pxssaddcellNum    b                   export
     DxssaddcellNum    Pi
     d value                          8F   VALUE
     D style                          3u 0 VALUE  OPTIONS(*NOPASS)
     D idx                            5u 0 VALUE  OPTIONS(*NOPASS)
     D merge                          3u 0 VALUE  OPTIONS(*NOPASS)
      /free
        if %parms = 1;
          Addcell();
        elseif %parms = 2;
          Addcell(style);
        elseif %parms = 3;
          Addcell(style:idx);
        elseif %parms = 4;
          Addcell(style:idx:merge);
        ENDIF;
        data += '<Data ss:Type="Number">' + %char(value)
               	+ '</Data></Cell>';
        addEntry();
      /end-free
     p                 e
     pxssaddcellDat    b                   export
     DxssaddcellDat    Pi
     D DateN                          8  0 VALUE
     D format                         5    value
     D style                          3u 0 VALUE  OPTIONS(*NOPASS)
     D idx                            5u 0 VALUE  OPTIONS(*NOPASS)
     D merge                          3u 0 VALUE  OPTIONS(*NOPASS)

     d aDate           s               d
     d datcar          s             10
      /free
        if %parms < 3;
          Addcell();
        elseif %parms =3;
          Addcell(style);
        elseif %parms = 4;
          Addcell(style:idx);
        elseif %parms = 5;
          Addcell(style:idx:merge);
        ENDIF;
        Select;
           When daten = 0;
              data += '<Data ss:Type="DateTime"></Data></Cell>';
              addEntry();
              return;
           When format = F_CYMD;
             Test(ed) *cymd daten;
             If not %error;
               aDate = %Date(%Dec(%Subst(
                         %EditC(daten:'X'):2:7):7:0):*cymd);
             ElseIf %rem(daten:100) = 99;
               aDate = D'9999-12-31';
             EndIf;
           When format = F_YMD;
             Test(ed) *ymd daten;
             If not %error;
               aDate = %Date(%Dec(%Subst(
                         %EditC(daten:'X'):3:6):6:0):*ymd);
             ElseIf %rem(daten:100) = 99;
               aDate = D'9999-12-31';
             EndIf;
           When format = F_DMY;
             Test(ed) *dmy daten;
             If not %error;
               aDate = %Date(%Dec(%Subst(
                         %EditC(daten:'X'):3:6):6:0):*dmy);
             ElseIf daten = 999999;
               aDate = D'9999-12-31';
             EndIf;
           When format = F_ISO;
             Test(ed) *iso daten;
             If not %error;
               aDate = %Date(daten:*iso);
             ElseIf %rem(daten:100) = 99;
               aDate = D'9999-12-31';
             EndIf;
           When format = F_MDY;
             Test(ed) *MDY daten;
             If not %error;
               aDate = %Date(%Dec(%Subst(
                         %EditC(daten:'X'):3:6):6:0):*mdy);
             ElseIf daten = 999999;
               aDate = D'9999-12-31';
             EndIf;
         EndSl;
         datcar = %char(aDate:*iso);
         data += '<Data ss:Type="DateTime">' + datcar
                + 'T00:00:00.000</Data></Cell>';
         addEntry();

      /end-free
     p                 e
     p XssGetRowId     B                   export
     D XssGetRowId     Pi             6
      /free
         return 'R' + %char(rowId);
      /end-free
     p                 e
     P XssGetColId     B                   export
     D XssGetColId     pi             5
      /free
         return 'C' + %char(colId);
      /end-free
     p                 e
     P XssGetCellId    B                   export
     D XssGetCellId    pi            12
      /free
         return %trim(XssGetRowId()) + XssGetColId();
      /end-free
     p                 e

     p XssSetAutoFilter...
     p                 b                   export
     D XssSetAutoFilter...
     D                 pi
     D  r1                           12    Value
     D  r2                           12    value
      /free
         autofilter = *on;
         range1 = r1;
         range2 = r2;
      /end-free
     p                 e
     p SetAutoFilter   b                   export
     D SetAutoFilter   pi
      /free
         data = '<AutoFilter x:Range="' + %trim(range1) + ':' + %trim(range2)
               + '" '+ EOR
               + 'xmlns="urn:schemas-microsoft-com:office:excel">' + EOR
               + '</AutoFilter>';
         addEntry();
      /end-free
     p                 e

     p XssFreezePane   b                   export
     D XssFreezePane   Pi
     d  Row                           5i 0 value
     d  Col                           5i 0 value
     d                                     OPTIONS(*NOPASS)
       FreezePane = *on;
       if row > 0;
         FreezeH = %editc(row:'Z');
       endif;
       if %parms = 2 and col > 0;
         FreezeV = %editc(Col:'Z');
       endif;
     p XssFreezePane   e

     p SetFreezePane   b
         data = '<WorksheetOptions xmlns='
               + '"urn:schemas-microsoft-com:office:excel">'   + EOR
               + '<Selected/>' + EOR
               + '<FreezePanes/>' + EOR
               + '<FrozenNoSplit/>' + EOR ;
         addEntry();
         if FreezeH <> ' ';
           data = '<SplitHorizontal>' + Freezeh +
                  '</SplitHorizontal>' +eor +
                  '<TopRowBottomPane>' + Freezeh +
                  '</TopRowBottomPane>' +eor  ;
           addEntry();
         endif;
         if FreezeV <> ' ';
           data = '<SplitVertical>' + Freezeh +
                  '</SplitVertical>' +eor +
                  '<LeftColumnRightPane>' + Freezeh +
                       '</LeftColumnRightPane>' +eor  ;
            addEntry();
          endif;
          data = '<ActivePane>2</ActivePane>' + EOR
               + '</WorksheetOptions> ' + EOR;
          addEntry();
     p SetFreezePane   e

     pXssclosefile     b                   export
     DXssclosefile     Pi
      /free
         If OpenRow;
           data = '</Row>';
           addEntry();
         ENDIF;
         if OpenSheet;
           data = '</Table>';
           AddEntry();
            if Autofilter;
             SetAutoFilter();
           ENDIF;
           if FreezePane;
             SetFreezePane();
           ENDIF;
           data = '</Worksheet>';
           AddEntry();
         ENDIF;
         data = '</Workbook>';
         addEntry();
         txtclofile();
      /end-free
     p                 e

      //--------------------------------------
       // Procedure name: addentry
       // Purpose: add a new line in IFS file
       //--------------------------------------
     P addentry        B
     D addEntry        PI

      /FREE

          data2 = %trimr(data) + EOR;
          txtwrite(%addr(data2):%len(%trimr(data2)));
      /end-free
     P addentry        e

     p init            b
     d init            pi
      /free
        reset openRow;
        reset openStyle;
        reset openSheet;
        reset style;
        idx = 0;
        HashUcs = hash;
        HashCar = HashUcs;
      /end-free
     p init            e
     p chgvar          b
     D chgvar          pi          2000
     D  var                           1    value
     D  input                         4    value
     D  string                     1024    value

     D  pos            s              5i 0
     D  newval         s             10    varying
     D  ret            s           2000
      /free
         newval = '&' + hashcar + 'x' + input + ';' ;
         ret = string;
         pos = %scan(var:ret);
         dow  pos > 0;
           ret = %replace(newval:ret :pos:1) ;
           pos = %scan(var:ret:pos + 1);
         enddo;
         return ret;
      /end-free
     p                 e
     p chgvar2         b
     D chgvar2         pi          2000
     D  var                           1    value
     D  input                         5    value
     D  string                     1024    value

     D  pos            s              5i 0
     D  newval         s             10    varying
     D  ret            s           2000
      /free
         newval = %trim(input) ;
         ret = %scanrpl(var:newval:string);
         return ret;
      /end-free
     p                 e
** ctdata(aa)
À00C0
Á00C1
Â00C2
Ã00C3
Ä00C4
Å00C5
Æ00C6
Ç00C7
È00C8
É00C9
Ê00CA
Ë00CB
Ì00CC
Í00CD
Î00CE
Ï00CF
Ð00D0
Ñ00D1
Ò00D2
Ó00D3
Ô00D4
Õ00D5
Ö00D6
#00A3
Ø00D8
Ù00D9
Ú00DA
Û00DB
Ü00DC
Ý00DD
Þ00DE
ß00DF
@00E0
á00E1
â00E2
ã00E3
ä00E4
å00E5
æ00E6
\00E7
}00E8
{00E9
ê00EA
ë00EB
ì00EC
í00ED
î00EE
ï00EF
ð00F0
ñ00F1
ò00F2
ó00F3
ô00F4
õ00F5
ö00F6
÷00F7
ø00F8
¦00F9
ú00FA
û00FB
ü00FC
ý00FD
þ00FE
ÿ00FF
