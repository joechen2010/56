			var nowurl = document.location.toString();
			//<![CDATA[
			var adLeftSrc = "images/left.swf"
			var adLeftFlash = "flash"
			var adLeftHref = ""
			var adLeftWidth = 100
			var adLeftHeight = 360

			var adRightSrc = "images/right.swf"
			var adRightFlash = "flash"
			var adRightHref = ""
			var adRightWidth = 100
			var adRightHeight = 360
			var marginTop = 90
			var marginLeft = 10
			var navUserAgent = navigator.userAgent
			function load(){
				judge();
				move();
			}
			function move() {
				judge();
				setTimeout("move();",80)
			}
			function judge(){
				if (navUserAgent.indexOf("Firefox") >= 0 || navUserAgent.indexOf("Opera") >= 0) {
					if (adLeftSrc != "") {document.getElementById("adLeftFloat").style.top = (document.body.scrollTop?document.body.scrollTop:document.documentElement.scrollTop) + ((document.body.clientHeight > document.documentElement.clientHeight)?document.documentElement.clientHeight:document.body.clientHeight) - adLeftHeight - marginTop + 'px';}
					if (adRightSrc != "") {
						document.getElementById("adRightFloat").style.top = (document.body.scrollTop?document.body.scrollTop:document.documentElement.scrollTop) + ((document.body.clientHeight > document.documentElement.clientHeight)?document.documentElement.clientHeight:document.body.clientHeight) - adRightHeight - marginTop + 'px';
						document.getElementById("adRightFloat").style.left = ((document.body.clientWidth > document.documentElement.clientWidth)?document.body.clientWidth:document.documentElement.clientWidth) - adRightWidth - marginLeft + 'px';
					}
				}
				else{
					if (adLeftSrc != "") {document.getElementById("adLeftFloat").style.top = (document.body.scrollTop?document.body.scrollTop:document.documentElement.scrollTop) + marginTop + 'px';}

					if (adRightSrc != "") {
						document.getElementById("adRightFloat").style.top = (document.body.scrollTop?document.body.scrollTop:document.documentElement.scrollTop) +  marginTop + 'px';
						document.getElementById("adRightFloat").style.left = ((document.documentElement.clientWidth == 0)?document.body.clientWidth:document.documentElement.clientWidth)  - adRightWidth - marginLeft + 'px';
					}
				}
				if (adLeftSrc != "") {document.getElementById("adLeftFloat").style.left = marginLeft + 'px';}
			}
			if (adLeftSrc != "") {
				if (adLeftFlash == "flash") {
					document.write("<div id=\"adLeftFloat\" style=\"position: absolute;width:" + adLeftWidth + ";\"><a href=\"" + adLeftHref +"\"><embed src=\"" + adLeftSrc + "\" quality=\"high\"  width=\"" + adLeftWidth + "\" height=\"" + adLeftHeight + "\" type=\"application/x-shockwave-flash\" id=\"sinadl\"></embed></a></div>");
				}
				else{
					document.write("<div id=\"adLeftFloat\" style=\"position: absolute;width:" + adLeftWidth + ";\"><a href=\"" + adLeftHref +"\"><img src=\"" + adLeftSrc + "\"  width=\"" + adLeftWidth + "\" height=\"" + adLeftHeight + "\"  border=\"0\" \></a></div>");
				}
			}
			if (adRightSrc != "") {
				if (adRightFlash == "flash") {
					document.write("<div id=\"adRightFloat\" style=\"position: absolute;width:" + adRightWidth + ";\"><a href=\"" + adRightHref +"\"><embed src=\"" + adRightSrc + "\" quality=\"high\"  width=\"" + adLeftWidth + "\" height=\"" + adRightHeight + "\" type=\"application/x-shockwave-flash\"  id=\"sinadl\"></a></embed></div>");
				}
				else{
					document.write("<div id=\"adRightFloat\" style=\"position: absolute;width:" + adRightWidth + ";\"><a href=\"" + adRightHref +"\"><img src=\"" + adRightSrc + "\"   width=\"" + adLeftWidth + "\" height=\"" + adRightHeight + "\"  border=\"0\"  \></a></div>");
				}
			}
			load();
			//]]>
