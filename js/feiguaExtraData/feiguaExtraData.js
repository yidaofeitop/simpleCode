// ==UserScript==
// @name         飞瓜数据商品页面数据
// @author          yidaofeitop
// @description  抽取飞官页面商品页面数据：商品名称、商品来源、商品链接……
// @match           *://dy.feigua.cn/Member*
// @require         https://cdn.bootcss.com/jquery/3.4.1/jquery.min.js
// @require         https://cdn.bootcss.com/xlsx/0.15.1/xlsx.full.min.js
// @version         0.0.1
// ==/UserScript==

(function() {
    'use strict';

    $(document).ready(function() {

        //debugger;

        //判断按钮是否存在，如不存在则添加如存在则无需处理
        //下载商品数据
        if($("#extraGoodsDataBtn").length==0){
            var $insertGoodsBtn=$("<button type='button'>下载商品数据</button>")
            $(".btns-area").before($insertGoodsBtn);
            $insertGoodsBtn.addClass("btn btn-primary");
            $insertGoodsBtn.attr("id","extraGoodsDataBtn");
            $insertGoodsBtn.bind("click",function(){extraGoodsData();});
        }
         //下载商品数据
        if($("#extraVideosDataBtn").length==0){
            var $insertVideosBtn=$("<button type='button'>下载视频数据</button>")
            $(".btns-area").before($insertVideosBtn);
            $insertVideosBtn.addClass("btn btn-primary");
            $insertVideosBtn.css("margin-left","20px")
            $insertVideosBtn.attr("id","extraVideosDataBtn");
            $insertVideosBtn.bind("click",function(){extraVideosData();});
        }
    });

    function extraVideosData(){
        var dataTrs=$("#js-blogger-history-awemes").find('tr');
        var trLength=dataTrs.length;
        if(trLength<1){
            alert("当前达人无视频");
            return;
        }
        var goodsNum=$("#AwemeCount_Data").text();
        if(trLength<goodsNum){
          alert("目前仅加载 "+trLength+" 个视频，视频总数量为：: "+goodsNum+" 请注意!");
        }

        //抽取每行的数据，构建一个json数据
        var index=0;
        var commodityArray=new Array();
        //构建表头部分
        var commodity=["标题","上传时间","点赞数","评论量","转发数","视频链接"];
        commodityArray.push(commodity);
        commodity=null;

        while(index<trLength){
            var simpleTr=dataTrs[index];
            commodity=new Array();
            //获取标题、话题、@、上线实践
            var videoTitle=$(simpleTr).find("td:eq(0)").find("div.item-title").find("a").text().replace(/ /g,'');
            var publishTime=$(simpleTr).find("td:eq(0)").find("div.item-times").text().match(/\d{4}[-\d{2}]*[\s][\d{2}\:]*/);
            var likeCommentForwardTd=$(simpleTr).find("td:eq(1)").find("div.v-icon-set-box").html().match(/\d+/g);
            var videoHref=$(simpleTr).find("td:eq(2)").find("div.mp-article-source").find("a:eq(2)").attr("href");


            //处理点赞……
            var like=likeCommentForwardTd[0];
            var comment=likeCommentForwardTd[1];
            var forward=likeCommentForwardTd[2];

            //处理数据
            commodity=new Array();
            commodity.push(videoTitle);
            commodity.push(publishTime);
            commodity.push(like);
            commodity.push(comment);
            commodity.push(forward);
            commodity.push(videoHref);

            //console.log(commodity);
            commodityArray.push(commodity);
            //增加序号
            index++;
        }
        console.log(commodityArray);

        //利用js-xlsx输出
        var sheet = XLSX.utils.aoa_to_sheet(commodityArray);
        openDownloadDialog(sheet2blob(sheet), '视频数据.xlsx');
        return commodityArray;


        return commodityArray;
    }


    function extraGoodsData(){
        var dataTrs=$("#table_goods").find('tr');
        var trLength=dataTrs.length;
        if(trLength<1){
            alert("当前抖音达人无带货数据或数据未加载完成");
            return;
        }
        var goodsNum=$(".js-thead").find("th:eq(0)").html().replace(/[^0-9]/ig,"");
        if(trLength<goodsNum){
          alert("目前仅加载 "+trLength+" 已上架商品数: "+goodsNum+" 请注意!");
        }

        //抽取每行的数据，构建一个json数据
        var index=0;
        var commodityArray=new Array();
        //构建表头部分
        var commodity=["商品名称"," 播主点赞增量","视频浏览量","订单数","单价","链接"];
        commodityArray.push(commodity);
        commodity=null;

        while(index<trLength){
            var simpleTr=dataTrs[index];
            commodity=new Array();
            //获取渠道名称
            var $titleATag=$(simpleTr).find("div[class='item-title js-goods-title']").find("a");
            var href=$titleATag.attr("href");
            var title=$titleATag.text();
            var like=$(simpleTr).find("td:eq(1)").text();
            var pageviews=$(simpleTr).find("td:eq(3)").text();
            var sale=$(simpleTr).find("td:eq(4)").text();
            var price=$(simpleTr).find("td:eq(5)").text();

            //处理数据
            commodity=new Array();
            commodity.push(title);
            commodity.push(like);
            commodity.push(pageviews.replace(/[^\d]/g,''));
            commodity.push(sale.replace(/[^\d]/g,''));
            commodity.push(href);
            commodity.push(price);

            //console.log(commodity);
            commodityArray.push(commodity);
            //增加序号
            index++;
        }
        console.log(commodityArray);

        //利用js-xlsx输出
        var sheet = XLSX.utils.aoa_to_sheet(commodityArray);
        openDownloadDialog(sheet2blob(sheet), '商品数据.xlsx');
        return commodityArray;


        return commodityArray;
    }


// 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
function sheet2blob(sheet, sheetName) {
  sheetName = sheetName || 'sheet1';
  var workbook = {
    SheetNames: [sheetName],
    Sheets: {}
  };
  workbook.Sheets[sheetName] = sheet;
  // 生成excel的配置项
  var wopts = {
    bookType: 'xlsx', // 要生成的文件类型
    bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
    type: 'binary'
  };
  var wbout = XLSX.write(workbook, wopts);
  var blob = new Blob([s2ab(wbout)], {type:"application/octet-stream"});
  // 字符串转ArrayBuffer
  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }
  return blob;
}
/**
 * 通用的打开下载对话框方法，没有测试过具体兼容性
 * @param url 下载地址，也可以是一个blob对象，必选
 * @param saveName 保存文件名，可选
 */
function openDownloadDialog(url, saveName)
{
  if(typeof url == 'object' && url instanceof Blob)
  {
    url = URL.createObjectURL(url); // 创建blob地址
  }
  var aLink = document.createElement('a');
  aLink.href = url;
  aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
  var event;
  if(window.MouseEvent) event = new MouseEvent('click');
  else
  {
    event = document.createEvent('MouseEvents');
    event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
  }
  aLink.dispatchEvent(event);
}


})();