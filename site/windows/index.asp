<html>
<head><%
select case month(Now)
case 1
mes = "Janeiro"
case 2
mes = "Fevereiro"
case 3
mes = "Março"
case 4
mes = "Abril"
case 5
mes = "Maio"
case 6
mes = "Junho"
case 7
mes = "Julho"
case 8
mes = "Agosto"
case 9
mes = "Setembro"
case 10
mes = "Outubro"
case 11
mes = "Novembro"
case 12
mes = "Dezembro"
end select
response.write "<title>Cat&aacute;logo Hermes de "& mes &"</title>"%>
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
<script language="JavaScript" type="text/javascript">
flippingBookCatalogoHermes = new FlippingBook();
flippingBookCatalogoHermes.pages = [ 
"paginas/01.jpg|",
"paginas/02.jpg|",
"paginas/03.jpg|",
"paginas/04.jpg|",
"paginas/05.jpg|",
"paginas/06.jpg|",
"paginas/07.jpg|",
"paginas/08.jpg|",
"paginas/09.jpg|",
"paginas/10.jpg|",
"paginas/11.jpg|",
"paginas/12.jpg|",
"paginas/13.jpg|",
"paginas/14.jpg|",
"paginas/15.jpg|",
"paginas/16.jpg|",
"paginas/17.jpg|",
"paginas/18.jpg|",
"paginas/19.jpg|",
"paginas/20.jpg|",
"paginas/21.jpg|",
"paginas/22.jpg|",
"paginas/23.jpg|",
"paginas/24.jpg|",
"paginas/25.jpg|",
"paginas/26.jpg|",
"paginas/27.jpg|",
"paginas/28.jpg|",
"paginas/29.jpg|",
"paginas/30.jpg|",
"paginas/31.jpg|",
"paginas/32.jpg|",
"paginas/33.jpg|",
"paginas/34.jpg|",
"paginas/35.jpg|",
"paginas/36.jpg|",
"paginas/37.jpg|",
"paginas/38.jpg|",
"paginas/39.jpg|",
"paginas/40.jpg|",
"paginas/41.jpg|",
"paginas/42.jpg|",
"paginas/43.jpg|",
"paginas/44.jpg|",
"paginas/45.jpg|",
"paginas/46.jpg|",
"paginas/47.jpg|",
"paginas/48.jpg|",
"paginas/49.jpg|",
"paginas/50.jpg|",
"paginas/51.jpg|",
"paginas/52.jpg|",
"paginas/53.jpg|",
"paginas/54.jpg|",
"paginas/55.jpg|",
"paginas/56.jpg|",
"paginas/57.jpg|",
"paginas/58.jpg|",
"paginas/59.jpg|",
"paginas/60.jpg|",
"paginas/61.jpg|",
"paginas/62.jpg|",
"paginas/63.jpg|",
"paginas/64.jpg|",
"paginas/65.jpg|",
"paginas/66.jpg|",
"paginas/67.jpg|",
"paginas/68.jpg|",
"paginas/69.jpg|",
"paginas/70.jpg|",
"paginas/71.jpg|",
"paginas/72.jpg|",
"paginas/73.jpg|",
"paginas/74.jpg|",
"paginas/75.jpg|",
"paginas/76.jpg|",
"paginas/77.jpg|",
"paginas/78.jpg|",
"paginas/79.jpg|",
"paginas/80.jpg|",
"paginas/81.jpg|",
"paginas/82.jpg|",
"paginas/83.jpg|",
"paginas/84.jpg|",
"paginas/85.jpg|",
"paginas/86.jpg|",
"paginas/87.jpg|",
"paginas/88.jpg|",
"paginas/89.jpg|",
"paginas/90.jpg|",
"paginas/91.jpg|",
"paginas/92.jpg|",
"paginas/93.jpg|",
"paginas/94.jpg|",
"paginas/95.jpg|",
"paginas/96.jpg|",
"paginas/97.jpg|",
"paginas/98.jpg|",
"paginas/99.jpg|",
"paginas/100.jpg|",
"paginas/101.jpg|",
"paginas/102.jpg|",
"paginas/103.jpg|",
"paginas/104.jpg|",
"paginas/105.jpg|",
"paginas/106.jpg|",
"paginas/107.jpg|",
"paginas/108.jpg|",
"paginas/109.jpg|",
"paginas/110.jpg|",
"paginas/111.jpg|",
"paginas/112.jpg|",
"paginas/113.jpg|",
"paginas/114.jpg|",
"paginas/115.jpg|",
"paginas/116.jpg|",
"paginas/117.jpg|",
"paginas/118.jpg|",
"paginas/119.jpg|",
"paginas/120.jpg|",
"paginas/121.jpg|",
"paginas/122.jpg|",
"paginas/123.jpg|",
"paginas/124.jpg|",
"paginas/125.jpg|",
"paginas/126.jpg|",
"paginas/127.jpg|",
"paginas/128.jpg|",
"paginas/129.jpg|",
"paginas/130.jpg|",
"paginas/131.jpg|",
"paginas/132.jpg|",
"paginas/133.jpg|",
"paginas/134.jpg|",
"paginas/135.jpg|",
"paginas/136.jpg|",
"paginas/137.jpg|",
"paginas/138.jpg|",
"paginas/139.jpg|",
"paginas/140.jpg|",
"paginas/141.jpg|",
"paginas/142.jpg|",
"paginas/143.jpg|",
"paginas/144.jpg|",
"paginas/145.jpg|",
"paginas/146.jpg|",
"paginas/147.jpg|",
"paginas/148.jpg|",
"paginas/149.jpg|",
"paginas/150.jpg|",
"paginas/151.jpg|",
"paginas/152.jpg|",
"paginas/153.jpg|",
"paginas/154.jpg|",
"paginas/155.jpg|",
"paginas/156.jpg|",
];

flippingBookCatalogoHermes.enlargedImages = [
"paginas/01.jpg|",
"paginas/02.jpg|",
"paginas/03.jpg|",
"paginas/04.jpg|",
"paginas/05.jpg|",
"paginas/06.jpg|",
"paginas/07.jpg|",
"paginas/08.jpg|",
"paginas/09.jpg|",
"paginas/10.jpg|",
"paginas/11.jpg|",
"paginas/12.jpg|",
"paginas/13.jpg|",
"paginas/14.jpg|",
"paginas/15.jpg|",
"paginas/16.jpg|",
"paginas/17.jpg|",
"paginas/18.jpg|",
"paginas/19.jpg|",
"paginas/20.jpg|",
"paginas/21.jpg|",
"paginas/22.jpg|",
"paginas/23.jpg|",
"paginas/24.jpg|",
"paginas/25.jpg|",
"paginas/26.jpg|",
"paginas/27.jpg|",
"paginas/28.jpg|",
"paginas/29.jpg|",
"paginas/30.jpg|",
"paginas/31.jpg|",
"paginas/32.jpg|",
"paginas/33.jpg|",
"paginas/34.jpg|",
"paginas/35.jpg|",
"paginas/36.jpg|",
"paginas/37.jpg|",
"paginas/38.jpg|",
"paginas/39.jpg|",
"paginas/40.jpg|",
"paginas/41.jpg|",
"paginas/42.jpg|",
"paginas/43.jpg|",
"paginas/44.jpg|",
"paginas/45.jpg|",
"paginas/46.jpg|",
"paginas/47.jpg|",
"paginas/48.jpg|",
"paginas/49.jpg|",
"paginas/50.jpg|",
"paginas/51.jpg|",
"paginas/52.jpg|",
"paginas/53.jpg|",
"paginas/54.jpg|",
"paginas/55.jpg|",
"paginas/56.jpg|",
"paginas/57.jpg|",
"paginas/58.jpg|",
"paginas/59.jpg|",
"paginas/60.jpg|",
"paginas/61.jpg|",
"paginas/62.jpg|",
"paginas/63.jpg|",
"paginas/64.jpg|",
"paginas/65.jpg|",
"paginas/66.jpg|",
"paginas/67.jpg|",
"paginas/68.jpg|",
"paginas/69.jpg|",
"paginas/70.jpg|",
"paginas/71.jpg|",
"paginas/72.jpg|",
"paginas/73.jpg|",
"paginas/74.jpg|",
"paginas/75.jpg|",
"paginas/76.jpg|",
"paginas/77.jpg|",
"paginas/78.jpg|",
"paginas/79.jpg|",
"paginas/80.jpg|",
"paginas/81.jpg|",
"paginas/82.jpg|",
"paginas/83.jpg|",
"paginas/84.jpg|",
"paginas/85.jpg|",
"paginas/86.jpg|",
"paginas/87.jpg|",
"paginas/88.jpg|",
"paginas/89.jpg|",
"paginas/90.jpg|",
"paginas/91.jpg|",
"paginas/92.jpg|",
"paginas/93.jpg|",
"paginas/94.jpg|",
"paginas/95.jpg|",
"paginas/96.jpg|",
"paginas/97.jpg|",
"paginas/98.jpg|",
"paginas/99.jpg|",
"paginas/100.jpg|",
"paginas/101.jpg|",
"paginas/102.jpg|",
"paginas/103.jpg|",
"paginas/104.jpg|",
"paginas/105.jpg|",
"paginas/106.jpg|",
"paginas/107.jpg|",
"paginas/108.jpg|",
"paginas/109.jpg|",
"paginas/110.jpg|",
"paginas/111.jpg|",
"paginas/112.jpg|",
"paginas/113.jpg|",
"paginas/114.jpg|",
"paginas/115.jpg|",
"paginas/116.jpg|",
"paginas/117.jpg|",
"paginas/118.jpg|",
"paginas/119.jpg|",
"paginas/120.jpg|",
"paginas/121.jpg|",
"paginas/122.jpg|",
"paginas/123.jpg|",
"paginas/124.jpg|",
"paginas/125.jpg|",
"paginas/126.jpg|",
"paginas/127.jpg|",
"paginas/128.jpg|",
"paginas/129.jpg|",
"paginas/130.jpg|",
"paginas/131.jpg|",
"paginas/132.jpg|",
"paginas/133.jpg|",
"paginas/134.jpg|",
"paginas/135.jpg|",
"paginas/136.jpg|",
"paginas/137.jpg|",
"paginas/138.jpg|",
"paginas/139.jpg|",
"paginas/140.jpg|",
"paginas/141.jpg|",
"paginas/142.jpg|",
"paginas/143.jpg|",
"paginas/144.jpg|",
"paginas/145.jpg|",
"paginas/146.jpg|",
"paginas/147.jpg|",
"paginas/148.jpg|",
"paginas/149.jpg|",
"paginas/150.jpg|",
"paginas/151.jpg|",
"paginas/152.jpg|",
"paginas/153.jpg|",
"paginas/154.jpg|",
"paginas/155.jpg|",
"paginas/156.jpg|",
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