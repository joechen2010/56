var myScroller1 = new Scroller(0, 0, 220, 92,1, 3);
myScroller1.setColors("#FFffff", "", "");
myScroller1.setFont("MS UI Gothic,888888,Osaka", 2);
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>商城公告</b><br>乐购网上商城正式开张了，欢迎收藏本站，欢迎光临本站挑选您喜爱的商品。</font>")
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>商城公告：</b><br>欢迎免费注册为会员，无论注册与否均可购物，但会员购物更加便捷。</font>")
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>商城公告：</b><br>开业酬宾，100个品种7折优惠，欢迎选购。</font>")
myScroller1.setPause(3000);
function runmikescroll() {
var layer;
var mikex, mikey;
 layer = getLayer("placeholder");
  mikex = getPageLeft(layer);
  mikey = getPageTop(layer);

  myScroller1.create();
  myScroller1.hide();
  myScroller1.moveTo(mikex, mikey);
  myScroller1.setzIndex(100);
  myScroller1.show();
}
window.onload=runmikescroll