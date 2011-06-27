function PhotoFilter(){
	this.CurrPos				= 0;
	this._ImgSrcList		= null;
	this._TimeList			= null;
	this._TitleList			= null;
	this._HrefList			= null;
	this._ObjAdContainer	= null;
	this._ObjText				= null;
	this._Speed					= 1000;
	this.ObjName				= null;
	this.Width					= 0;
	this.Height					= 0;
	this.AdFrame				= null;
	
	this.ArrImg					= null;
	this.ArrTime				= null;
	this.ArrTitle				= null;
	this.ArrLink				= null;
	
	this.theTimer				= null;
	
	this.ImgSrcList	= function(strImgSrcList){
		this._ImgSrcList = strImgSrcList;
		this.ArrImg = this._ImgSrcList.split('|||');
		//alert(this.ArrImg.length);
	}
	
	this.Init = function(){
		//this.AdFrame = document.frames['FlashAdFrame'];
		//this.AdFrame.document.write('<html><head></head><body></body></html>');
		//this.AdFrame.document.domain = '.job1998.com';
		//µ¼º½±í¸ñ
		var ObjTab = document.createElement("table");
		ObjTab.id = 'tabPhotoFilterMain';
		ObjTab.cellSpacing = 1;
		ObjTab.cellPadding = 1;
		ObjTab.border = 0;
		ObjTab.height = 15;
		ObjTab.bgColor = '#DDDDDD';
		ObjTab.style.filter = 'progid:DXImageTransform.Microsoft.Alpha(opacity=100)';
		ObjTab.style.position = 'absolute';
		document.body.appendChild(ObjTab);
		var ObjRow = ObjTab.insertRow(0);
		for (var i=0; i<this.ArrImg.length; i++){
			var ObjCell = ObjRow.insertCell(i);
			ObjCell.id = 'tdPhotoFilter_' + (i+1);
			ObjCell.width = 13;
			ObjCell.align = 'center';
			ObjCell.valign = 'middle';
			ObjCell.bgColor = '#999999';
			//ObjCell.fontColor = '#FFFFFF';
			ObjCell.style.cursor = 'hand';
			ObjCell.style.fontFamily = 'Arial';
			ObjCell.attachEvent('onclick', ClickPhotoTd);
			ObjCell.innerHTML = '<span style="font-size:7pt;color:#FFFFFF">' + (i+1) + '</span>';
			/*
			var ObjFileName = new CFileName(this.ArrImg[i]);
			var ObjAd;
			if (ObjFileName.Extend == 'swf'){
				ObjAd = this.CreateFlash(this.AdFrame.window.document, i, this.ArrImg[i]);
			} else {
				ObjAd = this.CreateImg(this.AdFrame.window.document, i, this.ArrImg[i]);
			}
			this.AdFrame.window.document.body.appendChild(ObjAd);
			*/
		}
		var strHTML = '<A id="FlashAdHref" name="FlashAdHref" Href=\'javascript:void(0);\' target=\'_blank\'><img border="0" id=\'imgPhotoFilter\' name=\'imgPhotoFilter\' src=\'\' style="FILTER: revealTrans(duration=2,transition=20)" width="' + this.Width + '" height="' + this.Height + '" style="display:none" />';
    strHTML += '<object id=\'swfPhotoFilter\' name=\'swfPhotoFilter\' style="display:none" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="' + this.Width + '" height="' + this.Height + '">';
    strHTML += '<param name="movie" value="../js/pixviewer.swf">';
    strHTML += '<param name="quality" value="high">';
    strHTML += '<param name="wmode" value="transparent">';
    strHTML += '<embed src="../js/pixviewer.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="' + this.Width + '" height="' + this.Height + '"></embed></object>';
    strHTML += '</A>';
    this._ObjAdContainer.innerHTML = strHTML;
		this.SetNagivatePos();
	}
	
	this.CreateImg = function(Doc, intId, strSrc){
		var ObjDiv = Doc.createElement('div');
		ObjDiv.id = 'FlashAdContainer_' + intId;
		//ObjDiv.style.position = 'absolute';
		var objImg = Doc.createElement("img");
		objImg.id = 'FlashAd_' + intId;
		objImg.src = strSrc;
		objImg.style.filter = 'revealTrans(duration=2,transition=' + Math.floor(Math.random()*28) + ')';
		
		ObjDiv.appendChild(objImg);
		//objImg.style.display = 'none';
		return ObjDiv;
	}
	
	this.CreateFlash = function(Doc, intId, strSrc){
		var ObjDiv = Doc.createElement('div');
		ObjDiv.id = 'FlashAdContainer_' + intId;
		//ObjDiv.style.position = 'absolute';
		var objImg = Doc.createElement("div");
		objImg.id = 'FlashAd_' + intId;
		strFlash = '<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="' + this.Width + '" height="' + this.Height + '">';
		strFlash +=	'<param name="movie" value="../index.files/%27%20%2B%20strSrc%20%2B%20%27">';
		strFlash +=	'<param name="quality" value="high">';
		strFlash +=	'<param name="wmode" value="transparent">';
		strFlash +=	'<embed src="../index.files/%27%20%2B%20strSrc%20%2B%20%27" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="' + this.Width + '" height="' + this.Height + '"></embed></object>';
		//alert(strFlash);
		objImg.innerHTML = strFlash;
		ObjDiv.appendChild(objImg);
		//ObjDiv.style.zIndex = 99;
		return ObjDiv;
	}
	
	this.TimeList	= function(strTimeList){
		this._TimeList = strTimeList;
		this.ArrTime = this._TimeList.split('|||');
	}
	
	this.TitleList	= function(strTitleList){
		this._TitleList = strTitleList;
		this.ArrTitle = this._TitleList.split('|||');
	}
	
	this.HrefList	= function(strHrefList){
		this._HrefList = strHrefList;
		this.ArrLink = this._HrefList.split('|||');
	}
	
	this.ObjAdContainer	= function(TheAdContaner){
		this._ObjAdContainer = TheAdContaner;
	}
	
	this.SetNagivatePos = function(){
		var ObjClient = GetClientRect(this._ObjAdContainer);
		var ObjTab = document.all.tabPhotoFilterMain;
		//alert(ObjTab.clientWidth);
		ObjTab.style.top = ObjClient.Top+ObjClient.Height-ObjTab.height;
		ObjTab.style.left = ObjClient.Left+ObjClient.Width-ObjTab.clientWidth+1;
		ObjTab.style.zIndex = 100;
	}
	
	this.ObjText = function(TheText){
		this._ObjText = TheText;
	}
	
	this.Speed = function(intSpeed){
		this._Speed = parseInt(intSpeed);
	}
	
	this.SetFilter = function(Obj){
		if (Obj.tagName.toLowerCase() == 'img'){
			Obj.filters.revealTrans.Transition = Math.floor(Math.random()*28);
			Obj.filters.revealTrans.apply();
			//alert(Obj.filters.revealTrans.Transition);
		}
	}
	
	this.PlayFilter = function(Obj){
		if (Obj.tagName.toLowerCase() == 'img'){
			Obj.filters.revealTrans.play();
		}
	}
	
	this.NextPhoto = function(){
		this.ProcessFilter();
		if(this.CurrPos < this.ArrImg.length-1){
			this.CurrPos++;
		}
		else {
			this.CurrPos = 0;
		}
	}
	
	this.ProcessFilter = function(){
		//var Obj = eval('this.AdFrame.window.document.all.FlashAdContainer_' + this.CurrPos);
		//if (Obj){
			for (var i=0; i<this.ArrImg.length; i++){
				var tdObj = eval('document.all.tdPhotoFilter_' + (i+1));
				tdObj.bgColor = '#000000';
			}
			var tdObj = eval('document.all.tdPhotoFilter_' + (this.CurrPos+1));
			tdObj.bgColor = '#FF0000';
			
			var ObjFileName = new CFileName(this.ArrImg[this.CurrPos]);
			var ObjAdSwf = document.all.swfPhotoFilter;
			var ObjAdImg = document.all.imgPhotoFilter;
			if (ObjFileName.Extend == 'swf'){
				ObjAdSwf.movie = this.ArrImg[this.CurrPos];
				ObjAdImg.style.display = 'none';
				ObjAdSwf.style.display = '';
			} else {
				this.SetFilter(ObjAdImg);
				ObjAdImg.src = this.ArrImg[this.CurrPos];
				ObjAdSwf.style.display = 'none';
				ObjAdImg.style.display = '';
				this.PlayFilter(ObjAdImg);
			}
			this._ObjText.innerHTML = '<A Href="' + this.ArrLink[this.CurrPos] + '" Target="_blank">' + this.ArrTitle[this.CurrPos] + '</A>';
			document.all.FlashAdHref.href = this.ArrLink[this.CurrPos];
			this.theTimer = setTimeout(this.ObjName + '.NextPhoto()', this.ArrTime[this.CurrPos]);
		//}
	}
	
	this.GoToTarget = function(intPos, Obj){
		if (this.theTimer){
			clearTimeout(this.theTimer);
		}
		var ObjTheAd = this._ObjAdContainer.firstChild;
		//alert(this.CurrPos);
		if (ObjTheAd.tagName){
			if (ObjTheAd.tagName.toLowerCase() == 'img'){
				ObjTheAd.filters.revealTrans.stop();
			}
		}
		this.CurrPos = parseInt(intPos)-1;
		this.NextPhoto();
	}
}>