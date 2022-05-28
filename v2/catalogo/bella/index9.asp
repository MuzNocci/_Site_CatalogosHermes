<html>
<head><title>Ponto Vai Ponto Vem</title>
<style type="text/css">
 	html, body, div, iframe { margin:0; padding:0; height:100%; } 
   	iframe { display:block; width:100%; border:none; } 
</style>
<script type="text/javascript" src="js/swfobject.js"></script>
<script type="text/javascript" src="js/flippingbook.js"></script>
<script type="text/javascript" src="js/jquery-1.4.2.min.js"></script>
</head>
<body style="margin:0px;">
<div id="fbContainer_CatalogoHermes">
	<div id="altmsg">
    	</div>
</div>
<%
spt = "<script language=""JavaScript"" type=""text/javascript"">"&VBnewline
spt = spt & "flippingBookCatalogoHermes = new FlippingBook();"&VBnewline
spt = spt & "flippingBookCatalogoHermes.pages = ["&VBnewline 
response.write spt

Function verificaImagem(url)
       
        On Error Resume Next
        
        Set objHttp = Server.CreateObject("MSXML2.XMLHTTP")
        objHttp.Open "GET", url, False
        objHttp.Send
        resultado = objHttp.getResponseHeader("Content-type")
        Set objHttp = Nothing
        
        If ((inStr(resultado,"image") > 0) And (Err.Number = 0)) Then
                resultado = """paginas/"&right(url,7)&"""|,"&vbnewline	
        Else
                resultado = ""
        End If
        
        verificaImagem = resultado
End Function

for i = 0 to 30
if len(i) = 1 then
u = "00"&i
end if
if len(i) = 2 then
u = "0"&i
end if
Response.Write verificaImagem("http://catalogobella.megamidiagroup.com.br/pvpv1/images_g/"&u&".jpg")
next

response.write "];"
%>

flippingBookCatalogoHermes.enlargedImages = [
"paginas/zoom-01.jpg|",
"paginas/zoom-02.jpg|",
"paginas/zoom-03.jpg|",
"paginas/zoom-04.jpg|",
"paginas/zoom-05.jpg|",
"paginas/zoom-06.jpg|",
"paginas/zoom-07.jpg|",
"paginas/zoom-08.jpg|",
"paginas/zoom-09.jpg|",
"paginas/zoom-10.jpg|",
"paginas/zoom-11.jpg|",
"paginas/zoom-12.jpg|",
"paginas/zoom-13.jpg|",
"paginas/zoom-14.jpg|",
"paginas/zoom-15.jpg|",
"paginas/zoom-16.jpg",
];

//flippingBookCatalogoHermes.pageLinks = ["|",""];

flippingBookCatalogoHermes.settings.uniqueSuffix = "CatalogoHermes";
flippingBookCatalogoHermes.stageWidth = "800"; 
flippingBookCatalogoHermes.stageHeight = "580";
flippingBookCatalogoHermes.settings.direction = "LTR";
flippingBookCatalogoHermes.settings.bookWidth = "666";
flippingBookCatalogoHermes.settings.bookHeight = "520";
flippingBookCatalogoHermes.settings.firstPageNumber = "1";
flippingBookCatalogoHermes.settings.navigationBar = "swf/navigation.swf";
flippingBookCatalogoHermes.settings.navigationBarPlacement = "bottom";
flippingBookCatalogoHermes.settings.pageBackgroundColor = 0xEEEEEE;
flippingBookCatalogoHermes.settings.backgroundColor = "dedede";
flippingBookCatalogoHermes.settings.backgroundImage = "img/background.jpg";
flippingBookCatalogoHermes.settings.backgroundImagePlacement = "fit";
flippingBookCatalogoHermes.settings.staticShadowsType = "Asymmetric";
flippingBookCatalogoHermes.settings.staticShadowsDepth = "1";
flippingBookCatalogoHermes.settings.autoFlipSize = "75";
flippingBookCatalogoHermes.settings.centerBook = true;
flippingBookCatalogoHermes.settings.scaleContent = true;
flippingBookCatalogoHermes.settings.alwaysOpened = false;
flippingBookCatalogoHermes.settings.flipCornerStyle = "manually";
flippingBookCatalogoHermes.settings.hardcover = false;
flippingBookCatalogoHermes.settings.downloadURL = "";
flippingBookCatalogoHermes.settings.downloadTitle = "";
flippingBookCatalogoHermes.settings.downloadSize = "";
flippingBookCatalogoHermes.settings.allowPagesUnload = false;
flippingBookCatalogoHermes.settings.fullscreenEnabled = true;
flippingBookCatalogoHermes.settings.zoomEnabled = true;
flippingBookCatalogoHermes.settings.zoomImageWidth = "740"; 
flippingBookCatalogoHermes.settings.zoomImageHeight = "1156"; 
flippingBookCatalogoHermes.settings.zoomUIColor = 0x8f9ea6;
flippingBookCatalogoHermes.settings.slideshowButton = false;
flippingBookCatalogoHermes.settings.slideshowAutoPlay = false;
flippingBookCatalogoHermes.settings.slideshowDisplayDuration = "5000";
flippingBookCatalogoHermes.settings.goToPageField = true;
flippingBookCatalogoHermes.settings.firstLastButtons = true;
flippingBookCatalogoHermes.settings.printEnabled = false;
// Bot? Imprimir
flippingBookCatalogoHermes.settings.zoomingMethod = "flash";
flippingBookCatalogoHermes.settings.soundControlButton = true;
flippingBookCatalogoHermes.settings.showUnderlyingPages = false;
flippingBookCatalogoHermes.settings.fullscreenHint = "";
flippingBookCatalogoHermes.settings.zoomHintEnabled = "true";
flippingBookCatalogoHermes.settings.zoomOnClick = true;
flippingBookCatalogoHermes.settings.moveSpeed = "2";
flippingBookCatalogoHermes.settings.closeSpeed = "3";
flippingBookCatalogoHermes.settings.gotoSpeed = "3";
flippingBookCatalogoHermes.settings.rigidPageSpeed = "5";
flippingBookCatalogoHermes.settings.zoomHint = "Duplo clique para aumentar";
flippingBookCatalogoHermes.settings.printTitle = "";
flippingBookCatalogoHermes.settings.downloadComplete = "";
flippingBookCatalogoHermes.settings.dropShadowEnabled = true;
flippingBookCatalogoHermes.settings.flipSound = "mp3/flip.mp3";
flippingBookCatalogoHermes.settings.hardcoverSound = "mp3/hardcover.mp3";
flippingBookCatalogoHermes.settings.preloaderType = "Progress Bar";
flippingBookCatalogoHermes.settings.Ioader = true;
flippingBookCatalogoHermes.settings.frameColor = 0xFFFFFF;
flippingBookCatalogoHermes.settings.frameWidth = "0";
flippingBookCatalogoHermes.containerId = "fbContainer_CatalogoHermes";
flippingBookCatalogoHermes.create("swf/flippingbook.swf");
jQuery.noConflict();
</script>
</body>
</html>