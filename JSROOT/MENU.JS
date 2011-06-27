var myScroller1 = new Scroller(0, 0, 220, 92,1, 3);
myScroller1.setColors("#FFffff", "","");
myScroller1.setFont("MS UI Gothic,888888,Osaka", 2);
myScroller1.addItem("<b>网站公告-N.e.w.s</b>");
myScroller1.addItem("12312-3123-123<br><font color=000000>我们很好</font>");
myScroller1.setPause(500);
function runmikescroll() {
var layer;var mikex, mikey;
 layer = getLayer("placeholder");
  mikex = getPageLeft(layer);
  mikey = getPageTop(layer);
  myScroller1.create();
  myScroller1.hide();
  myScroller1.moveTo(mikex, mikey);
  myScroller1.setzIndex(100);
  myScroller1.show();
