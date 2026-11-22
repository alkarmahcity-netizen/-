const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,Table,TableRow,TableCell,WidthType,AlignmentType,HeadingLevel,BorderStyle,ShadingType,LevelFormat,Footer,Header,PageNumber,TableLayoutType}=require('docx');
const F="Arial";
const r=(t,o={})=>new TextRun({text:t,font:{ascii:F,hAnsi:F,cs:F},rightToLeft:true,size:o.size||24,bold:o.b,color:o.color,italics:o.i});
const ltr=(t,o={})=>new TextRun({text:"\u202A"+t+"\u202C",font:{ascii:"Consolas",hAnsi:"Consolas",cs:F},size:o.size||22,bold:o.b,rightToLeft:false});
// parts: string or ['code',text] or ['b',text]
const runs=parts=>(Array.isArray(parts)?parts:[parts]).map(p=>typeof p==="string"?r(p):p[0]==="code"?ltr(p[1]):p[0]==="b"?r(p[1],{b:true}):r(p[1],{i:true,color:"555555"}));
const P=(parts,o={})=>new Paragraph({bidirectional:true,alignment:o.center?AlignmentType.CENTER:AlignmentType.RIGHT,spacing:{after:o.after??120,before:o.before??0,line:300},children:runs(parts),numbering:o.bullet?{reference:"b",level:0}:o.num?{reference:o.num,level:0}:undefined,keepNext:o.keepNext});
const H1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,bidirectional:true,alignment:AlignmentType.RIGHT,children:[r(t,{size:30,b:true,color:"1D3557"})],keepNext:true,border:{bottom:{style:BorderStyle.SINGLE,size:6,color:"A57F35",space:4}}});
const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,bidirectional:true,alignment:AlignmentType.RIGHT,children:[r(t,{size:26,b:true,color:"1D3557"})],keepNext:true});
const B=parts=>P(parts,{bullet:true,after:80});
let numN=0;const steps=[];const N=(list)=>{const ref="n"+(numN++);steps.push(ref);return list.map(x=>P(x,{num:ref,after:80}))};
const W=9638;
function T(rows,widths,head=true){const cw=widths.map(w=>Math.round(W*w));cw[cw.length-1]=W-cw.slice(0,-1).reduce((a,b)=>a+b,0);
  return new Table({visuallyRightToLeft:true,width:{size:W,type:WidthType.DXA},columnWidths:cw,layout:TableLayoutType.FIXED,
   rows:rows.map((row,i)=>new TableRow({tableHeader:head&&i===0,cantSplit:true,children:row.map((c,j)=>new TableCell({width:{size:cw[j],type:WidthType.DXA},margins:{top:80,bottom:80,left:100,right:100},
     shading:head&&i===0?{type:ShadingType.CLEAR,fill:"1D3557",color:"auto"}:i%2===0?{type:ShadingType.CLEAR,fill:"F2F4F7",color:"auto"}:undefined,
     children:[new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,children:head&&i===0?[r(String(c),{b:true,color:"FFFFFF",size:22})]:runs(c)})]}))}))})}
const sp=()=>new Paragraph({children:[],spacing:{after:80}});
const box=(parts)=>new Paragraph({bidirectional:true,alignment:AlignmentType.RIGHT,spacing:{after:160,before:80,line:300},shading:{type:ShadingType.CLEAR,fill:"FBF5E8",color:"auto"},border:{right:{style:BorderStyle.SINGLE,size:18,color:"A57F35",space:6}},children:runs(parts)});

const body=[
 new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:60},children:[r("محافظة الأنبار — ديوان المحافظة",{b:true,size:26})]}),
 new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:240},children:[r("قسم الشؤون المالية · قسم الشؤون الهندسية",{size:22,color:"555555"})]}),
 new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:120},children:[r("دليل ربط برنامج الشؤون المالية بنظام الشؤون الهندسية الموحد عبر شبكة الديوان",{b:true,size:36,color:"1D3557"})]}),
 new Paragraph({bidirectional:true,alignment:AlignmentType.CENTER,spacing:{after:360},children:[r("دليل تنفيذي لمسؤول الشبكة ومستخدمي البرنامجين · تشرين الثاني 2026",{size:22,color:"555555"})]}),

 H1("1. الغرض والصورة العامة"),
 P("يقع قسم الشؤون المالية على بعد نحو 300 متر من قسم الشؤون الهندسية، ولكل منهما راوتر. يشرح هذا الدليل كيف تُجمع الحاسبتان في شبكة واحدة، ثم كيف يقرأ كل برنامج بيانات الآخر تلقائياً من مجلد مشترك، دون إنترنت ودون خادم."),
 T([["المكوّن","حاسبة القسم الهندسي","حاسبة القسم المالي"],
    ["البرنامج","نظام الشؤون الهندسية الموحد","برنامج الشؤون المالية (AnbarFinance، نسخة ويندوز)"],
    ["الملف الذي يُشارَك",[["code","registry.json"]],[["code","anbar-data.json"]]],
    ["مكانه",[["code","%LOCALAPPDATA%\\AnbarProjects"]],[["code","المستندات\\AnbarFinance"]]],
    ["اسم المشاركة",[["code","AnbarProjects"]],[["code","AnbarFinance"]]],
    ["العنوان المقترح",[["code","192.168.1.50"]],[["code","192.168.1.60"]]]],[0.22,0.39,0.39]),
 sp(),
 P([["b","مبدأ الربط: "],"كل برنامج يكتب في ملفه على حاسبته، ويقرأ ملف البرنامج الآخر عبر الشبكة قراءة فقط. لذلك لا يمكن أن يفسد أحدهما بيانات الآخر."]),
 box([["b","تنبيه: "],"العناوين في هذا الدليل أمثلة بنطاق ",["code","192.168.1.x"],". استبدلها بنطاق شبكة المحافظة الفعلي، واختر عناوين غير مستعملة وخارج المدى الذي يوزعه الراوتر تلقائياً."]),

 H1("2. ما يلزم"),
 T([["المادة","العدد","ملاحظة"],
    ["جهاز جسر لاسلكي خارجي 5GHz (Point-to-Point)، مثل TP-Link CPE510 أو Ubiquiti LiteBeam","2","واحد على سطح كل مبنى، متقابلان بلا عوائق (أشجار أو أبنية)"],
    ["كابل شبكة Cat6 خارجي مع موصلات","حسب المسافة","من الجهاز على السطح إلى الراوتر أو السويتش، أقل من 100 متر لكل جهة"],
    ["حامل أو عمود تثبيت","2","بارتفاع يضمن خط نظر مباشر بين الجهازين"],
    ["محوّل الطاقة PoE","2","يأتي غالباً مع الجهاز"],
    ["حساب مستخدم محلي للمشاركة","1 في كل حاسبة","مثال: anbar-share بكلمة سر قوية"]],[0.55,0.15,0.30]),
 sp(),
 P([["i","البديل عن الجسر اللاسلكي: كابل ألياف ضوئية بين المبنيين مع محوّلَي وسائط (Media Converter). كابل الشبكة النحاسي العادي لا يصلح لأن حدّه نحو 100 متر."]]),

 H1("3. الخطوة الأولى: ربط المبنيين بالجسر اللاسلكي"),
 ...N([
  "ثبّت الجهازين على سطحي المبنيين بحيث يرى كل منهما الآخر بخط نظر مباشر، ووجّههما نحو بعضهما.",
  "صل كل جهاز بمحوّل PoE، ومن المحوّل كابلاً إلى الراوتر أو السويتش في المبنى نفسه.",
  "ادخل إلى صفحة إعداد الجهاز الأول من المتصفح (العنوان مكتوب على الجهاز أو في كتيّبه)، واضبطه نقطة وصول (Access Point) بوضع الجسر (Bridge أو Transparent).",
  "اضبط الجهاز الثاني محطة (Station أو Client) بوضع الجسر، واجعله يتصل بالأول.",
  "فعّل التشفير WPA2 بكلمة سر قوية، وأخفِ اسم الشبكة اللاسلكية إن أمكن.",
  [["b","أعطِ كل جهاز عنوان إدارة ثابتاً "],"من نطاق الشبكة، مثل ",["code","192.168.1.250"]," و",["code","192.168.1.251"],"، وأوقف فيه توزيع العناوين (DHCP)."],
  "راقب قوة الإشارة في صفحة الإعداد، وعدّل التوجيه حتى تصل إلى أفضل قيمة."]),
 box([["b","ملاحظة: "],"أسماء القوائم تختلف بين الشركات المصنّعة. المطلوب في النهاية أن يعمل الجهازان كجسر شفاف ينقل الشبكة كما هي، لا كراوتر ثانٍ."]),

 H1("4. الخطوة الثانية: جعل الراوتر الثاني جزءاً من الشبكة نفسها"),
 P("هذه أهم خطوة. إذا بقي الراوتر الثاني يعمل راوتراً مستقلاً فلن ترى الحاسبتان بعضهما."),
 ...N([
  "ادخل إلى صفحة إعداد الراوتر الثاني (في المبنى الأبعد عن الراوتر الرئيسي).",
  [["b","أوقف خدمة DHCP "],"فيه، حتى يأخذ كل جهاز عنوانه من الراوتر الرئيسي."],
  ["أعطِه عنواناً ثابتاً من نطاق الشبكة الرئيسية، مثل ",["code","192.168.1.2"],"."],
  [["b","صل كابل الجسر بأحد منافذ LAN في الراوتر، وليس منفذ WAN أو Internet."]],
  "إن كان في الراوتر خيار «وضع نقطة الوصول (Access Point Mode)» فاختره، فهو ينفذ الخطوات السابقة دفعة واحدة."]),
 box([["b","الخطأ الأكثر شيوعاً: "],"توصيل الراوتر الثاني من منفذ WAN. عندها تصبح خلفه شبكة منفصلة بعناوين مختلفة، فيفشل الربط حتى لو كان الإنترنت يعمل."]),

 H1("5. الخطوة الثالثة: عنوان ثابت لكل حاسبة"),
 P("العنوان الثابت ضروري، لأن البرنامجين يحفظان مسار الملف بعنوان الحاسبة الأخرى."),
 ...N([
  "في كل حاسبة: الإعدادات ← الشبكة والإنترنت ← Ethernet (أو Wi-Fi) ← تعيين IP ← تحرير ← يدوي ← IPv4.",
  [["b","حاسبة القسم الهندسي: "],"العنوان ",["code","192.168.1.50"],"، القناع ",["code","255.255.255.0"],"، البوابة ",["code","192.168.1.1"],"."],
  [["b","حاسبة القسم المالي: "],"العنوان ",["code","192.168.1.60"],"، القناع ",["code","255.255.255.0"],"، البوابة ",["code","192.168.1.1"],"."],
  ["خادم DNS: عنوان الراوتر الرئيسي ",["code","192.168.1.1"],"، أو ",["code","8.8.8.8"],"."]]),

 H1("6. الخطوة الرابعة: اختبار الاتصال"),
 ...N([
  "في حاسبة القسم المالي افتح «موجّه الأوامر» (اكتب cmd في قائمة ابدأ).",
  ["اكتب: ",["code","ping 192.168.1.50"]," ثم Enter. يجب أن تظهر أسطر «Reply from»."],
  ["كرر العكس من حاسبة القسم الهندسي: ",["code","ping 192.168.1.60"],"."]]),
 P("إذا لم يظهر رد فراجع جدول حل المشكلات في القسم 12 قبل المتابعة."),

 H1("7. الخطوة الخامسة: تهيئة ويندوز للمشاركة (في الحاسبتين)"),
 ...N([
  "الإعدادات ← الشبكة والإنترنت ← خصائص الشبكة ← نوع الشبكة: خاصة (Private).",
  "لوحة التحكم ← مركز الشبكة والمشاركة ← تغيير إعدادات المشاركة المتقدمة ← في «خاصة»: تشغيل اكتشاف الشبكة، وتشغيل مشاركة الملفات والطابعات.",
  [["b","أنشئ حساب مستخدم محلياً "],"باسم ",["code","anbar-share"]," وكلمة سر قوية: الإعدادات ← الحسابات ← مستخدمون آخرون ← إضافة حساب. أنشئ الحساب نفسه بكلمة السر نفسها في الحاسبتين."],
  "تأكد أن جدار حماية ويندوز يسمح بـ«مشاركة الملفات والطابعات» على الشبكة الخاصة، وهو مسموح تلقائياً عادة بعد الخطوة 2."]),

 H1("8. الخطوة السادسة: مشاركة مجلد الموقف الهندسي"),
 P([["b","على حاسبة القسم الهندسي فقط:"]]),
 ...N([
  ["افتح مستكشف الملفات، واكتب في شريط العنوان ",["code","%LOCALAPPDATA%"]," ثم Enter."],
  ["انقر بالزر الأيمن على المجلد ",["code","AnbarProjects"]," ← خصائص ← تبويب «مشاركة» ← مشاركة متقدمة."],
  ["فعّل «مشاركة هذا المجلد» واجعل اسم المشاركة ",["code","AnbarProjects"],"."],
  ["الأذونات: احذف «الجميع (Everyone)»، وأضف ",["code","anbar-share"]," بصلاحية «قراءة» فقط."],
  ["في تبويب «الأمان» للمجلد نفسه: أضف ",["code","anbar-share"]," بصلاحية «قراءة وتنفيذ» و«قراءة»."]]),

 H1("9. الخطوة السابعة: مشاركة مجلد بيانات المالية"),
 P([["b","على حاسبة القسم المالي فقط:"]]),
 ...N([
  ["افتح مجلد «المستندات»، ثم انقر بالزر الأيمن على المجلد ",["code","AnbarFinance"]," ← خصائص ← مشاركة ← مشاركة متقدمة."],
  ["اسم المشاركة ",["code","AnbarFinance"],"، والأذونات: ",["code","anbar-share"]," بصلاحية «قراءة» فقط، مع حذف «الجميع»."],
  "في تبويب «الأمان» أضف المستخدم نفسه بصلاحية «قراءة»."]),
 box([["b","إذا كان ملف بيانات المالية في مكان آخر: "],"في برنامج المالية افتح قائمة «مساعدة» ← «مكان ملف البيانات» لمعرفة مساره، وشارك المجلد الذي فيه."]),

 H1("10. الخطوة الثامنة: اختبار المجلدين من مستكشف الملفات"),
 ...N([
  ["من حاسبة القسم المالي اكتب في شريط العنوان: ",["code","\\\\192.168.1.50\\AnbarProjects"]],
  ["عند طلب كلمة السر أدخل المستخدم ",["code","anbar-share"]," وكلمة سره، واختر «تذكّر بيانات الاعتماد»."],
  ["يجب أن يظهر الملف ",["code","registry.json"],"."],
  ["من حاسبة القسم الهندسي اكتب: ",["code","\\\\192.168.1.60\\AnbarFinance"]," ويجب أن يظهر ",["code","anbar-data.json"],"."]]),

 H1("11. الخطوة التاسعة: ربط البرنامجين"),
 H2("في برنامج الشؤون المالية (حاسبة القسم المالي)"),
 ...N([
  "قائمة «ملف» ← «اختيار ملف الموقف الهندسي (registry.json)…».",
  ["اكتب في خانة اسم الملف: ",["code","\\\\192.168.1.50\\AnbarProjects\\registry.json"]," ثم «فتح»."],
  "يعيد البرنامج تحميل نفسه. افتح تبويب «الربط الهندسي» وتأكد أنه يعرض مسار الملف وعدد المشاريع الهندسية والمشاريع المربوطة."]),
 H2("في نظام الشؤون الهندسية الموحد (حاسبة القسم الهندسي)"),
 ...N([
  "افتح «💰 ربط المالية» ← «📂 اختيار الملف يدوياً».",
  ["اكتب: ",["code","\\\\192.168.1.60\\AnbarFinance\\anbar-data.json"]," ثم «فتح»."],
  "افتح أي مشروع في الموقف العام وتأكد من ظهور قسم «💰 من برنامج الشؤون المالية» في بطاقته."]),
 P("بعدها يعمل التحديث تلقائياً: برنامج المالية يقرأ الموقف الهندسي كل 30 ثانية، والنظام الهندسي يتابع تغيّر ملف المالية."),

 H1("12. حل المشكلات"),
 T([["المشكلة","السبب المحتمل","الحل"],
  ["لا يرد ping","الراوتر الثاني موصول من WAN، أو الجسر يعمل كراوتر، أو جدار الحماية يمنع الرد","راجع الخطوتين 1 و2. اسمح بـ«File and Printer Sharing (Echo Request - ICMPv4-In)» في جدار الحماية"],
  ["الحاسبتان بعناوين من نطاقين مختلفين","DHCP ما زال يعمل في الراوتر الثاني","أوقف DHCP فيه، وأعد تشغيل الحاسبة"],
  ["يطلب كلمة السر في كل مرة","بيانات الاعتماد غير محفوظة","لوحة التحكم ← إدارة بيانات الاعتماد ← بيانات اعتماد Windows ← إضافة: العنوان والمستخدم وكلمة السر"],
  ["«تم رفض الوصول»","الأذونات غير مضبوطة","راجع تبويبي «مشاركة» و«الأمان» للمجلد (الخطوتان 6 و7)"],
  ["البرنامج: «لم يُعثر على الملف»","الحاسبة الأخرى مطفأة أو تغيّر عنوانها","شغّل الحاسبة الأخرى، وتأكد أن عنوانها ما زال ثابتاً"],
  ["الأرقام لا تتحدث","الملف مختار من مسار قديم أو نسخة احتياطية","أعد اختيار الملف بمسار الشبكة الصحيح (الخطوة 9)"]],[0.25,0.35,0.40]),

 H1("13. الأمان والنسخ الاحتياطي"),
 B("المشاركة قراءة فقط في الاتجاهين. لا تمنح صلاحية «تغيير» أو «تحكم كامل» لأي مستخدم على الشبكة."),
 B("لا تشارك مجلد البرنامج كاملاً ولا القرص كاملاً. شارك المجلدين المذكورين فقط."),
 B("كلمة سر المستخدم anbar-share مستقلة عن كلمات سر البرنامجين، وتُحفظ لدى مسؤول الشبكة."),
 B("برنامج المالية يحفظ نسخة احتياطية يومية تلقائياً في مجلد backups بجانب ملف البيانات. يُستحسن نسخها أسبوعياً إلى قرص خارجي."),
 B("لا تفتح ملفي البيانات يدوياً ولا تعدّلهما. التعديل يتم من البرنامجين فقط."),

 H1("14. إن تعذّر ربط المبنيين بشبكة"),
 B([["b","كابل ألياف ضوئية "],"بين المبنيين مع محوّلَي وسائط. أثبت من الجسر اللاسلكي، لكن مدّه أصعب."]),
 B([["b","مزامنة عبر الإنترنت: "],"تثبيت Google Drive للحاسوب أو OneDrive على الحاسبتين، ونسخ الملفين إلى مجلد متزامن. يعمل، لكن بتأخير دقائق ويحتاج إنترنت مستقراً."]),
 B([["b","رفع يدوي: "],"يرفع القسم الهندسي ملف registry.json دورياً في تبويب «الربط الهندسي» لصفحة النظام المالي على claude.ai."]),

 H1("15. قائمة التحقق بعد التنفيذ"),
 T([["البند","تم","التاريخ والملاحظة"],
  ["الجسر اللاسلكي يعمل وقوة الإشارة جيدة","☐",""],
  ["الراوتر الثاني بلا DHCP وموصول من LAN","☐",""],
  [["عنوان ثابت للهندسية ",["code","192.168.1.50"]," وللمالية ",["code","192.168.1.60"]],"☐",""],
  ["ping ناجح في الاتجاهين","☐",""],
  ["مشاركة AnbarProjects قراءة فقط","☐",""],
  ["مشاركة AnbarFinance قراءة فقط","☐",""],
  ["برنامج المالية يعرض الموقف الهندسي في تبويب «الربط الهندسي»","☐",""],
  ["النظام الهندسي يعرض بيانات المالية في بطاقة المشروع","☐",""]],[0.55,0.10,0.35]),
 sp(),
 T([["","مسؤول الشبكة","رئيس قسم الشؤون المالية","رئيس قسم الشؤون الهندسية"],
  ["الاسم","","",""],["التوقيع","","",""],["التاريخ","","",""]],[0.16,0.28,0.28,0.28]),
];

const doc=new Document({creator:"قسم الشؤون المالية",title:"دليل ربط البرنامجين عبر الشبكة",
 styles:{default:{document:{run:{font:F,size:24,rightToLeft:true}}},paragraphStyles:[
  {id:"Heading1",name:"Heading 1",basedOn:"Normal",next:"Normal",quickFormat:true,run:{font:F,size:30,bold:true,color:"1D3557"},paragraph:{spacing:{before:360,after:160},keepNext:true,outlineLevel:0}},
  {id:"Heading2",name:"Heading 2",basedOn:"Normal",next:"Normal",quickFormat:true,run:{font:F,size:26,bold:true,color:"1D3557"},paragraph:{spacing:{before:200,after:100},keepNext:true,outlineLevel:1}}]},
 numbering:{config:[{reference:"b",levels:[{level:0,format:LevelFormat.BULLET,text:"•",alignment:AlignmentType.RIGHT,style:{paragraph:{indent:{right:500,hanging:300}}}}]},
   ...steps.map(ref=>({reference:ref,levels:[{level:0,format:LevelFormat.DECIMAL,text:"%1.",alignment:AlignmentType.RIGHT,style:{paragraph:{indent:{right:500,hanging:360}}}}]}))]},
 sections:[{properties:{page:{margin:{top:1134,bottom:1134,left:1134,right:1134}}},
  headers:{default:new Header({children:[new Paragraph({bidirectional:true,alignment:AlignmentType.LEFT,children:[r("دليل ربط برنامج الشؤون المالية بنظام الشؤون الهندسية الموحد",{size:18,color:"888888"})]})]})},
  footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({children:[PageNumber.CURRENT," / ",PageNumber.TOTAL_PAGES],font:F,size:18,color:"888888"})]})]})},
  children:body}]});
Packer.toBuffer(doc).then(b=>{fs.writeFileSync(process.argv[2],b);console.log('ok')});
