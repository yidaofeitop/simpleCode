// ==UserScript==
// @icon            https://login.aliexpress.com/favicon.ico
// @name            阿里巴巴国际站物流模板导入程序 
// @author          yidaofeitop
// @description     从指定物流模板Excel导入数据到阿里巴巴国际站-第一稿（基础数据导入以及国家数据筛查，无任何验证）
// @match           *://freighttemplate.aliexpress.com/wsproduct/freight/editFreightTemplate.htm?spm=*&id=* 
// @require         https://cdn.bootcss.com/jquery/3.4.1/jquery.min.js
// @require         https://cdn.bootcss.com/xlsx/0.15.1/xlsx.full.min.js
// @version         0.0.1
// ==/UserScript==

(function () {
    'use strict'; 
  	
  
//构建基础
//按钮
var $uploadBtn=$('<label for="ShippingTemplate" style="position: relative;"><input type="button" style="background-color: rgb(0, 123, 255);  border-color:  #6C757D; border-style: solid; border-radius: 4px;border-width: 0.666667px; color: rgb(255, 255, 255); display: inline-block; font-size: 16px;font-weight: 400;line-height: 24px;margin:4px 0px ;padding: 6px 12px; text-align: center;  vertical-align: middle; 	" value="点我上传"><span>自定义运费模板上传</span><input type="file" style="position: absolute;left: 0;top: 0;opacity: 0;" class="shipPriceExcelSZ"></label>');
//物流对应的国家列表
var shippingKindCountryMap=new Map();
shippingKindCountryMap.set("freight-custom-setting-CN_SUPER_ECONOMY_G","DZ,AF,AL,AD,AO,AR,AM,AU,AT,AZ,BH,BD,BE,BZ,BJ,BT,BO,BA,BW,BG,BF,BI,CM,CA,CV,TD,CL,CO,CR,HR,CY,CZ,ZR,DK,DJ,EC,EG,GQ,EE,ET,FK,FO,FJ,FI,FR,PF,GA,GM,DE,GH,GI,GR,GU,GT,GW,GY,HN,HU,IS,IN,ID,IQ,IE,IL,IT,CI,JP,JO,KH,KZ,KE,KI,KW,KG,LA,LV,LB,LS,LR,LY,LI,LT,LU,MK,MG,MW,MY,MV,ML,MT,MR,MU,MX,MC,MN,MNE,MA,MZ,MM,NA,NP,NL,NC,NZ,NI,NE,NG,NO,OM,PK,PA,PG,PY,PE,PH,PL,PT,PR,QA,MD,RE,RO,RW,SM,SN,SRB,SC,SG,SK,SI,SB,ZA,KR,LK,SR,SZ,SE,CH,TJ,TH,TLS,TG,TO,TN,TR,TM,UG,UK,TZ,US,UY,UZ,VU,VE,VN,ZM,ZW,BN,GE,GN,SJ,SX,SV,TT,AN,MH,AS,VI,FM,PM,VG,GF,GD,AI,GGY,GP,WF,KY,KM,MP,CK,AG,LC");
shippingKindCountryMap.set("freight-custom-setting-CAINIAO_EXPEDITED_ECONOMY","KR,ES,RU");
shippingKindCountryMap.set("freight-custom-setting-SUPER_ECONOMY_SG","BZ,CR,CA,HN,NI,PA,GT,MX,PR,US,RU,ES,AF,AL,DZ,BJ,BF,TD,ZR,LR,ML,SN,TG,TN,AD,AO,BW,LS,MG,MW,NA,RW,SZ,AR,AZ,AM,BG,AU,AT,BH,BD,TR,BO,EC,IQ,LI,BT,BA,BI,FK,CM,CV,PE,CO,HR,CY,CZ,DK,DJ,SC,EG,GQ,GA,GY,SR,EE,ET,FO,FJ,FI,FR,PF,GM,GH,GI,IE,GR,GU,GW,HU,IS,IN,ID,IL,IT,CI,JP,KH,LA,KE,KI,TO,VU,KG,LV,LB,LY,LT,LU,MK,MY,MV,MT,MR,NE,MN,MNE,MZ,MM,NP,NL,NC,NG,NO,OM,PK,PG,PH,PL,QA,MD,SRB,RE,RO,SM,SG,SK,SI,SB,ZA,KR,LK,SE,CH,TJ,TH,TM,UK,TZ,UY,VE,VN,ZM,ZW,BN,MA,MC,KZ,UG,NZ,DE,PY,TLS,MU,PT,UZ,BE,JO,KW");
shippingKindCountryMap.set("freight-custom-setting-SGP_OMP","CA,AR,IE,AU,BR,BE,PL,DK,DE,FR,NL,CZ,US,MX,NO,PT,SE,TR,UA,ES,IL,IT,UK,CL,RU,AF,AL,DZ,AS,AD,AO,AI,AG,AM,AW,AT,AZ,BH,BD,BB,BY,BZ,BJ,BM,BT,BA,BW,BG,BF,BI,CV,KY,CF,TD,CX,CC,CO,KM,CG,CR,HR,CY,ZR,DJ,DO,EC,EG,SV,GQ,ER,EE,ET,FK,FO,FJ,FI,PF,GA,GM,GE,GH,GI,GR,GL,GD,GP,GN,GF,HT,HN,HU,IS,IN,ID,IQ,CI,JM,JP,JO,KH,KZ,KE,KI,KW,KG,LA,LV,LB,LS,LR,LY,LI,LT,LU,MK,MG,MW,MY,MV,ML,MT,MH,MQ,MR,MU,MC,MN,MS,MA,MZ,MM,NA,NR,NP,AN,NC,NZ,NI,NE,NG,NF,MP,OM,PK,PA,PG,PY,PE,PH,PR,QA,MD,RE,RO,RW,SH,KN,LC,PM,VC,WS,SM,ST,SN,SRB,SC,SL,SG,SK,SI,SO,ZA,KR,LK,SR,SZ,CH,TJ,TH,BS,VA,TP,TG,TO,TT,TN,TM,TC,TV,UG,TZ,UY,UZ,VU,VE,VN,VG,VI,WF,EH,ZM,ZW,BN,PN,PS,ASC,CK,DM");
shippingKindCountryMap.set("freight-custom-setting-SINOTRANS_PY","ES");
shippingKindCountryMap.set("freight-custom-setting-YANWEN_ECONOMY","AF,IE,EE,AT,AU,BE,IS,PL,DK,DE,FR,FI,NL,CA,CZ,HR,LV,LT,US,MD,MX,NO,PT,SE,CH,SK,SI,TH,TR,NZ,HU,IL,IT,IN,UK,CL,PR,KZ,CY,RO,PE,MT,BA,CO,CR,LK,ZA,NG,KE,GE,JP,KR,ID,AL,DZ,AR,OM,AZ,EG,ET,AD,AO,PG,PK,PY,BH,PA,BG,BJ,BW,BT,BF,BI,GQ,TL,TG,FO,PF,PH,FJ,CV,FK,GM,GU,GY,MNE,HN,KI,DJ,KG,GN,GW,GH,GA,KH,ZW,CM,QA,KW,LS,LA,LB,LR,RE,LU,RW,MG,MV,MW,ML,MK,MU,MR,MN,BD,MM,MA,MC,MZ,NA,NP,NI,NE,SRB,SN,SC,SM,SZ,SR,SB,TJ,TZ,TO,TN,TM,VU,GT,VE,BN,UG,UY,UZ,GR,CI,NC,AM,JO,VN,ZM,TD,GI,SJ,SX,SV,BO,BZ,TT,EC,AN,LI,MH,AS,VI,FM,PM,VI,GF,GD,AI,LC,GGY,GP,WF,KY,KM,MP,CK,AG");
shippingKindCountryMap.set("freight-custom-setting-SUNYOU_ECONOMY","BZ,AF,AL,DZ,BJ,BF,TD,ZR,LR,ML,SN,TG,TN,AD,AO,BW,LS,MG,MW,NA,RW,SZ,AR,AZ,AM,BG,AU,AT,BH,BD,TR,BO,CR,EC,IQ,LI,BT,BA,BI,FK,CM,CA,CV,PE,CO,HR,CY,CZ,DK,DJ,SC,EG,GQ,GA,GY,HN,NI,SR,EE,ET,FO,FJ,FI,FR,PF,GM,PA,GH,GI,IE,GR,GU,GT,GW,HU,IS,IN,ID,IL,IT,CI,JP,KH,LA,KE,KI,TO,VU,KG,LV,LB,LY,LT,LU,MK,MY,MV,MT,MR,NE,MX,MN,MNE,MZ,MM,NP,NL,NC,NG,NO,OM,PK,PG,PH,PL,PR,QA,MD,SRB,RE,RO,SM,SG,SK,SI,SB,ZA,KR,LK,SE,CH,TJ,TH,TM,UK,TZ,UY,VE,VN,ZM,ZW,BN,MA,MC,KZ,UG,NZ,DE,PY,TLS,MU,PT,US,UZ,BE,JO,KW");
shippingKindCountryMap.set("freight-custom-setting-SF_EPARCEL_OM","EE,FI,LT,NO,PL,LV,SE");

 
//插入按钮
var $shippingPriceTdTags=$('table[class="table-list logistic-list ui-table"]').find("tbody").find("tr").find("td:eq(1)");

$shippingPriceTdTags.append($uploadBtn);

$('input.shipPriceExcelSZ').change(function(e){
		var shippingTemplate=e.target.files;
    var shippingTemplateReader=new FileReader(); 
		
  	//获取对应的物流国家数据
    var $shipPriceDataStoreTag=$(this).parent("label").parent("td").find("textarea:first"); 
    var standardShipCountries=shippingKindCountryMap.get($shipPriceDataStoreTag.attr("name"));
  	//console.log(standardShipCountries);
  
    //读取 Excel 文件 
    shippingTemplateReader.onload=function(ev){ 
        var shippingTemplateDataArray=[]; 
        try{
            //以二进制读取整个表格，本次仅针对一个sheet的表格
            var shippingTemplateData = ev.target.result; 
            var workbook=XLSX.read(shippingTemplateData,{type:"binary"}); 

            
             for (var sheet in workbook.Sheets) {
                if (workbook.Sheets.hasOwnProperty(sheet)) { 
                    shippingTemplateDataArray = shippingTemplateDataArray.concat(XLSX.utils.sheet_to_formulae(workbook.Sheets[sheet]));
                }
            } 
        }catch(e){
            console.log('文件类型不正确');
            return;
        } 
        
        //处理表格数据，后续如有分支模板直接调用不同的方法
        var shippingPriceCountryMap=excelToJSON(shippingTemplateDataArray);
      	//console.log(shippingPriceCountryMap);
        //统一处理
        var shipWeightDefinesJSON=dealWtihShippingPriceCountryMap(shippingPriceCountryMap,standardShipCountries);
        //修改相关按钮状态(否则数据无法提交)，并输出到textarea 
      	$shipPriceDataStoreTag.parent("div").find("input.custom-logistic").attr("checked","checked");
      	//$shipPriceDataStoreTag.parent("div").parent("td").parent("tr").find("td:first").find('input[class="logistic-checkbox"]').attr("checked","checked");
        $shipPriceDataStoreTag.html(shipWeightDefinesJSON); 
        //console.log(shipWeightDefinesJSON);
        
      	console.log($shipPriceDataStoreTag.html());

    }   
    shippingTemplateReader.readAsBinaryString(shippingTemplate[0]);   
});

  
  
//运费模板
//@param 表格数据，以单元格排列形成的数组
//@return Map结构数据,Key：运费组合字符串，value：国家字符串
function excelToJSON(excleArrayData){
    var shippingPriceCountryMap=new Map();

    var arrayDataindex=1;//此值会与表格的表头行数有关 

    while(arrayDataindex<excleArrayData.length/10){
        //从表格中获取所需数据
        var country=excleArrayData[arrayDataindex*10+2];
        var firstWeight=excleArrayData[arrayDataindex*10+5];;
        var firstPrice=excleArrayData[arrayDataindex*10+6];
        var continuedWeightInterval=excleArrayData[arrayDataindex*10+7]
        var continuedWeightEnd=excleArrayData[arrayDataindex*10+8]
        var continuedPrice=excleArrayData[arrayDataindex*10+9];
        var continuedWeightStart=0.01;
        //预设固定值
        
        
        
        country=country.slice(-2);
        firstWeight=firstWeight.substr(firstWeight.indexOf("=")+1);
        continuedWeightStart=firstWeight;
        continuedWeightInterval=continuedWeightInterval.substr(continuedWeightInterval.indexOf("=")+1);
        continuedWeightEnd=continuedWeightEnd.substr(continuedWeightEnd.indexOf("=")+1); 
        firstPrice=parseFloat(firstPrice.substr(firstPrice.indexOf("=")+1)).toFixed(2);
        continuedPrice=parseFloat(continuedPrice.substr(continuedPrice.indexOf("=")+1)).toFixed(2);

        //处理价格为负数的情况，全部变为0
        if(firstPrice<=0)       {           firstPrice=0;       }
        if(continuedPrice<=0)   {           continuedPrice=0;   }


        //将数据存储至于Map中
        //key：首重+"|"+首重价格+"|"+续重开始+"|"+续重结束+"|"+续重步增+"|"+续重价格
        var keyShippingPrice=firstWeight+"|"+firstPrice+"|"+continuedWeightStart+"|"+continuedWeightInterval+"|"+continuedWeightEnd+"|"+continuedPrice;

        //已经存在
        if(shippingPriceCountryMap.has(keyShippingPrice)){
            var valueCountries=shippingPriceCountryMap.get(keyShippingPrice);
            valueCountries=valueCountries.concat(","+country); 
            shippingPriceCountryMap.set(keyShippingPrice,valueCountries);
        }else{
            shippingPriceCountryMap.set(keyShippingPrice,country);
        } 

        arrayDataindex++;
    }
    return shippingPriceCountryMap; 
}
  
  
/*
//功能：处理Map数据分解成对应的JSON数据，并剔除掉已经设置运费的国家
//@param shippingPriceCountryMap:Map结构数据，key为运费相关信息，value为地区信息
//@param standardShipCountries:标准运费地区，初始值各自设定
//@return 运费JSON数据
*/
function dealWtihShippingPriceCountryMap(shippingPriceCountryMap,standardShipCountries){

    var iterator = shippingPriceCountryMap[Symbol.iterator]();

    var valueCountries="", keyShippingPrice="";
    var defines=[]; 
    for (let item of iterator) {
        keyShippingPrice=item[0];
        valueCountries=item[1];
				
      	//分解valueCountries,剔除standardShipCountries中相同的字段
			  standardShipCountries=removeCountry(standardShipCountries,valueCountries);
       
        //分解key与value值
        var keyShippingPriceArray=keyShippingPrice.split("|");
        /*var keyShippingPrice=firstWeight+"|"+firstPrice+"|"+continuedWeightStart+"|"+continuedWeightInterval+"|"+continuedWeightEnd+"|"+continuedPrice;*/
        

        var country=valueCountries;
        var firstWeight=keyShippingPriceArray[0];
        var firstPrice=keyShippingPriceArray[1];
        var continuedWeightStart=keyShippingPriceArray[2];
        var continuedWeightInterval=keyShippingPriceArray[3];
        var continuedWeightEnd=keyShippingPriceArray[4];
        var continuedPrice=keyShippingPriceArray[5];

        
        //构建单个国家的define 
        var countryShipPrice={};
        countryShipPrice.addFreight='';
        
        countryShipPrice.defineWeights=[{"endWeight":firstWeight,"intervalPrice":firstPrice,"intervalWeight":firstWeight,"startWeight":"0" },{"endWeight":continuedWeightEnd, "intervalPrice":continuedPrice,"intervalWeight":continuedWeightInterval,"startWeight":continuedWeightStart}] 
        countryShipPrice.endOrderNum='';
        countryShipPrice.minFreight='';
        countryShipPrice.perAddNum='';
     
        countryShipPrice.shippingCountry=country;
        countryShipPrice.startOrderNum='';           
        countryShipPrice.type="weight";
        
        defines.push(countryShipPrice);  
    } 

    var shipWeightDefines={}
    shipWeightDefines.defines=defines; 
    shipWeightDefines.freeShipCountry="";
    shipWeightDefines.notShippedCountries=""; 
  	shipWeightDefines.standardShipCountry=standardShipCountries; 
    shipWeightDefines.standardShipDiscount= "0";
    shipWeightDefines.standards=[];
    
  	//console.log("dealWtihShippingPriceCountryMap 输出数据:")
  	//console.log(JSON.stringify(shipWeightDefines)); 
  	
    return JSON.stringify(shipWeightDefines); 
}

  /*
 	//去除标准运费中的某些国家模板 
 	//@param standardShipCountries:标准运费地区，初始值参考阿里巴巴
 	//@param removeCoutries:剔除的地区字符串，以「,」隔开
 	//@return 标准运费地区字符串
 	*/
  function removeCountry(standardShipCountries, removeCoutries){

    if(!(standardShipCountries!=null&&standardShipCountries.length>0)){
      return "";
    }
    //将removeCoutries分解成单个国家简写
    var removeCountryArray=removeCoutries.split(",");
    var removeCountry="";
    var removeCountryIndex=0;

    for (var index = 0; index<removeCountryArray.length; index++) {
      var removeCountry=removeCountryArray[index];

      removeCountryIndex=standardShipCountries.indexOf(removeCountry); 
      //删除掉相关的国家
      if(removeCountryIndex!=0){
        standardShipCountries=standardShipCountries.replace(","+removeCountry,"");
      }else{
        standardShipCountries=standardShipCountries.replace(removeCountry+",","");
      } 
      //判断国家填写项目是否还有其他的
      if(standardShipCountries!=null&&standardShipCountries.length==0){
        break;
      }

    } 
    return standardShipCountries;
  } 
 
})();