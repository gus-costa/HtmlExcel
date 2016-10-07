<?php

class HtmlExcel {
    private $root = 'C:/E64B5C52';
    private $folder = 'Html2Excel';
    private $path = 'C:/E64B5C52/Html2Excel';
    private $mainFile = 'spreadsheet.htm';

    private $spreadsheets = array();
    private $info = array();
    private $css;

    public function addSheet($name, $contents){
        $this->info[] = $name;
        $this->spreadsheets[] = $contents;
    }

    public function setCss($css)
    {
        $this->css = $css;
    }

    public function headers($name = "spreadsheet.xls"){
        header("Content-type: application/vnd.ms-excel; charset=UTF-8");
        header("Content-type: application/force-download");
        header("Content-Disposition: attachment; filename=".$name);
    }

    public function buildFile(){
        $count = count($this->info);
        if ($count == 0) return '';
        elseif ($count == 1) return $this->buildSingleSheet();
        else return $this->buildMultiSheet();
    }

    private function buildMultiSheet(){
        $boundary = "----=_NextPart_01D21572.46A0BD00";
        $parts = array();
        $parts[] = 'MIME-Version: 1.0
X-Document-Type: Workbook
Content-Type: multipart/related; boundary="'.$boundary.'"';

        $parts[] = $this->buildMain();

        if (!empty($this->css)){
            $parts[] = $this->buildCss();
        }

        foreach ($this->info as $k => $v) {
            $parts[] = $this->buildSheet($k);
        }

        $parts[] = $this->buildFilelist();

        $ret = implode("\r\n\r\n--{$boundary}\r\n", $parts);
        $ret .= "\r\n--{$boundary}--\r\n";

        return $ret;
    }

    private function buildSingleSheet(){
            $ret='<html xmlns:x="urn:schemas-microsoft-com:office:excel">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <!--[if gte mso 9]>
    <xml>
        <x:ExcelWorkbook>
            <x:ExcelWorksheets>
                <x:ExcelWorksheet>
                    <x:Name><![CDATA['.substr($this->info[0], 0, 31).']]></x:Name>
                    <x:WorksheetOptions>
                        <x:Print>
                            <x:ValidPrinterInfo/>
                        </x:Print>
                    </x:WorksheetOptions>
                </x:ExcelWorksheet>
            </x:ExcelWorksheets>
        </x:ExcelWorkbook>
    </xml>
    <![endif]-->
    <style type="text/css">
'.$this->css.'
    </style>
</head>
<body>
';
            $ret.=$this->spreadsheets[0];
            $ret.='
</body>
</html>';
            return $ret;
    }

    private function buildMain(){
        $ret='Content-Location: file:///'.$this->root.'/'.$this->mainFile.'
Content-Transfer-Encoding: quoted-printable
Content-Type: text/html; charset="UTF-8"

<html xmlns:v=3D"urn:schemas-microsoft-com:vml"
xmlns:o=3D"urn:schemas-microsoft-com:office:office"
xmlns:x=3D"urn:schemas-microsoft-com:office:excel"
xmlns=3D"http://www.w3.org/TR/REC-html40">

<head>
<meta name=3D"Excel Workbook Frameset">
<meta http-equiv=3DContent-Type content=3D"text/html; charset=3Dutf-8">
<meta name=3DProgId content=3DExcel.Sheet>
<meta name=3DGenerator content=3D"Microsoft Excel 14">
<link rel=3DFile-List href=3D"'.$this->folder.'/filelist.xml">
';
        foreach ($this->info as $k=>$v) {
            $ret.='<link id=3D"shLink" href=3D"'.$this->folder.'/'.$k.'.htm">';
        }

        $ret.='
<link id=3D"shLink">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
';
        foreach ($this->info as $k => $v) {
            $ret.='   <x:ExcelWorksheet>
    <x:Name><![CDATA['.substr($v, 0, 31).']]></x:Name>
    <x:WorksheetSource HRef=3D"'.$this->folder.'/'.$k.'.htm"/>
   </x:ExcelWorksheet>
';
        }

        $ret.='  </x:ExcelWorksheets>';
        if (!empty($this->css))
            $ret.= '<x:Stylesheet HRef=3D"'.$this->folder.'/stylesheet.css"/>';
        $ret.='
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body>
</body>
</html>';

        return $ret;
    }

    private function buildSheet($key){
        $ret = 'Content-Location: file:///'.$this->path.'/'.$key.'.htm
Content-Transfer-Encoding: quoted-printable
Content-Type: text/html; charset="UTF-8"

<html xmlns:v=3D"urn:schemas-microsoft-com:vml"
xmlns:o=3D"urn:schemas-microsoft-com:office:office"
xmlns:x=3D"urn:schemas-microsoft-com:office:excel"
xmlns=3D"http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=3DContent-Type content=3D"text/html; charset=3Dutf-8">
<meta name=3DProgId content=3DExcel.Sheet>
<meta name=3DGenerator content=3D"Microsoft Excel 14">
<link id=3DMain-File rel=3DMain-File href=3D"../'.$this->mainFile.'">
<link rel=3DFile-List href=3Dfilelist.xml>
<![if IE]>
<base href=3D"file:///'.$this->path.'\\'.$key.'.htm"
id=3D"webarch_temp_base_tag">
<![endif]>
';
        if (!empty($this->css))
            $ret.='<link rel=3DStylesheet href=3Dstylesheet.css>';
        $ret.='
</head>

<body>';

        $ret.=str_replace("=", "=3D", $this->spreadsheets[$key]);

        $ret.='</body>
</html>';
        return $ret;
    }

    private function buildCss(){
        $ret='Content-Location: file:///'.$this->path.'/stylesheet.css
Content-Transfer-Encoding: quoted-printable
Content-Type: text/css; charset="utf-8"

';
        $ret .= $this->css;
        return $ret;
    }

    private function buildFilelist(){
        $ret = 'Content-Location: file:///'.$this->path.'/filelist.xml
Content-Transfer-Encoding: quoted-printable
Content-Type: text/xml; charset="utf-8"

<xml xmlns:o=3D"urn:schemas-microsoft-com:office:office">
 <o:MainFile HRef=3D"../'.$this->mainFile.'"/>';

        if (!empty($this->css)){
            $ret.= "\r\n";
            $ret.=' <o:File HRef=3D"stylesheet.css"/>';
        }

        foreach ($this->info as $k=>$v) {
            $ret.= "\r\n";
            $ret.=' <o:File HRef=3D"'.$k.'.htm"/>';
        }
        $ret .= "\r\n";

        $ret.=' <o:File HRef=3D"filelist.xml"/>
</xml>';
        return $ret;
    }

}