var myScroller1 = new Scroller(0, 0, 220, 92,1, 3);
myScroller1.setColors("#FFffff", "", "");
myScroller1.setFont("MS UI Gothic,888888,Osaka", 2);
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>�̳ǹ���</b><br>�ֹ������̳���ʽ�����ˣ���ӭ�ղر�վ����ӭ���ٱ�վ��ѡ��ϲ������Ʒ��</font>")
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>�̳ǹ��棺</b><br>��ӭ���ע��Ϊ��Ա������ע�������ɹ������Ա������ӱ�ݡ�</font>")
myScroller1.addItem("<font color=yellow style= font-size:10pt><b>�̳ǹ��棺</b><br>��ҵ�����100��Ʒ��7���Żݣ���ӭѡ����</font>")
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