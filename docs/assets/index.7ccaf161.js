var v=Object.defineProperty;var S=Object.getOwnPropertySymbols;var I=Object.prototype.hasOwnProperty,P=Object.prototype.propertyIsEnumerable;var x=(e,l,r)=>l in e?v(e,l,{enumerable:!0,configurable:!0,writable:!0,value:r}):e[l]=r,y=(e,l)=>{for(var r in l||(l={}))I.call(l,r)&&x(e,r,l[r]);if(S)for(var r of S(l))P.call(l,r)&&x(e,r,l[r]);return e};import{b as i,F as M,l as N,o as O,c as L,a as _,w as U,p as k,d as z,r as B,e as Y,f as D}from"./vendor.f32d8308.js";const Q=function(){const l=document.createElement("link").relList;if(l&&l.supports&&l.supports("modulepreload"))return;for(const t of document.querySelectorAll('link[rel="modulepreload"]'))o(t);new MutationObserver(t=>{for(const s of t)if(s.type==="childList")for(const c of s.addedNodes)c.tagName==="LINK"&&c.rel==="modulepreload"&&o(c)}).observe(document,{childList:!0,subtree:!0});function r(t){const s={};return t.integrity&&(s.integrity=t.integrity),t.referrerpolicy&&(s.referrerPolicy=t.referrerpolicy),t.crossorigin==="use-credentials"?s.credentials="include":t.crossorigin==="anonymous"?s.credentials="omit":s.credentials="same-origin",s}function o(t){if(t.ep)return;t.ep=!0;const s=r(t);fetch(t.href,s)}};Q();var n={NONE:0,OPEN_TAG:1<<1,CLOSE_TAG:1<<2,SELF_CLOSING:1<<3,ATTR_NAME:1<<4,ATTR_VALUE:1<<5,TEXT:1<<6,COMMENT:1<<7,ONELINE:1<<8,MULTILINE:1<<9,SINGLE_QUOTE:1<<10,DOUBLE_QUOTE:1<<11,STYLE_TAG:1<<12,SCRIPT_TAG:1<<13};const V=function(){var e=document.createElement("div");function l(r){return r&&typeof r=="string"&&(r=r.replace(/<script[^>]*>([\S\s]*?)<\/script>/gim,""),r=r.replace(/<\/?\w(?:[^"'>]|"[^"]*"|'[^']*')*>/gim,""),e.innerHTML=r,r=e.textContent,e.textContent=""),r}return l}();function j(e,l){var r=0,o=e[r],t=n.TEXT,s=n.NONE,c=n.NONE;function a(p,h){return p.length<=1&&!h?p===o:p===(h||e.substring(r-p.length,r))}function f(p,h){return e.substring(p,h)}function E(p){return p===t}function d(p){return p===s}var u=[0];u.set=function(p){u[u.length-1]=p},u.get=function(){return u[u.length-1]};for(var m=new function(){this[n.TEXT]={enter:function(){d(n.OPEN_TAG|n.CLOSE_TAG)||d(n.STYLE_TAG)||d(n.SCRIPT_TAG)||u.set(r)},exit:function(){var p=f(u.get(),r);!p||l.onText(p)},do:function(){a("<")&&(t=n.OPEN_TAG|n.CLOSE_TAG)}},this[n.OPEN_TAG|n.CLOSE_TAG]={enter:function(){u.push(r)},exit:function(){var p=u.pop();E(n.OPEN_TAG)&&u.set(p)},do:function(){a("!--")?t=n.COMMENT:a("/")?t=n.CLOSE_TAG:a(" ")?t=n.TEXT:!a("!")&&!a("-")&&(t=n.OPEN_TAG)}},this[n.OPEN_TAG]={exit:function(){var p=f(u.get(),r);!p||(p==="style"&&(c=n.STYLE_TAG),p==="script"&&(c=n.SCRIPT_TAG),l.onOpenTag(p))},do:function(){a(">")?t=n.SELF_CLOSING:a(" ")&&(t=n.ATTR_NAME)}},this[n.CLOSE_TAG]={enter:function(){u.set(r)},exit:function(){var p=f(u.get(),r);!p||l.onCloseTag(p)},do:function(){a(">")&&(t=n.TEXT)}},this[n.SELF_CLOSING]={enter:function(){u.push(r)},exit:function(){var p=u.pop();r=p-1,c&n.STYLE_TAG&&(t=n.OPEN_TAG|n.STYLE_TAG),c&n.SCRIPT_TAG&&(t=n.OPEN_TAG|n.SCRIPT_TAG)},do:function(){a("/>",f(r-2,r))?this.mark():this.close(),t=n.TEXT},mark:function(){l.onSelfClose()},close:function(){l.onOpenTagExit()}},this[n.ATTR_NAME]={enter:function(){u.set(r)},exit:function(){if(!a("/",f(r-1,r))){var p=f(u.get(),r);!p||l.onAttrName(p)}},do:function(){a(" ")||a("/")?this.repeat():a(">")?t=n.SELF_CLOSING:a("=")&&(t=n.ATTR_VALUE)},repeat:function(){this.exit(),u.set(r+1)}},this[n.ATTR_VALUE]={enter:function(){a("'")?(t|=n.SINGLE_QUOTE,u.set(r+1)):a('"')?(t|=n.DOUBLE_QUOTE,u.set(r+1)):u.set(r)},exit:function(){var p=f(u.get(r),r);!p||l.onAttrValue(p)},do:function(){a(" ")?t=n.ATTR_NAME:a(">")&&(t=n.SELF_CLOSING)}},this[n.ATTR_VALUE|n.SINGLE_QUOTE]={exit:this[n.ATTR_VALUE].exit,do:function(){a("'")&&(t=n.ATTR_NAME)}},this[n.ATTR_VALUE|n.DOUBLE_QUOTE]={exit:this[n.ATTR_VALUE].exit,do:function(){a('"')&&(t=n.ATTR_NAME)}},this[n.STYLE_TAG]={do:function(){a("/*")?t=n.STYLE_TAG|n.COMMENT:a("'")?t=n.STYLE_TAG|n.SINGLE_QUOTE:a('"')?t=n.STYLE_TAG|n.DOUBLE_QUOTE:a("<")&&a("</style>",f(r,r+8))&&(t=n.TEXT,r--,this.close())},close:function(){var p=u.pop();u.set(p),c^=n.STYLE_TAG}},this[n.OPEN_TAG|n.STYLE_TAG]={enter:function(){u.push(r)},do:this[n.STYLE_TAG].do},this[n.STYLE_TAG|n.COMMENT]={do:function(){a("*/")&&(t=n.STYLE_TAG)}},this[n.STYLE_TAG|n.SINGLE_QUOTE]={do:function(){a("'")&&(t=n.STYLE_TAG)}},this[n.STYLE_TAG|n.DOUBLE_QUOTE]={do:function(){a('"')&&(t=n.STYLE_TAG)}},this[n.SCRIPT_TAG]={do:function(){a("/*")?t=n.SCRIPT_TAG|n.COMMENT|n.MULTILINE:a("//")?t=n.SCRIPT_TAG|n.COMMENT|n.ONELINE:a("'")?t=n.SCRIPT_TAG|n.SINGLE_QUOTE:a('"')?t=n.SCRIPT_TAG|n.DOUBLE_QUOTE:a("`")?t=n.SCRIPT_TAG|n.MULTILINE:a("<")&&a("<\/script>",f(r,r+9))&&(t=n.TEXT,r--,this.close())},close:function(){var p=u.pop();u.set(p),c^=n.SCRIPT_TAG}},this[n.OPEN_TAG|n.SCRIPT_TAG]={enter:function(){u.push(r)},exit:this[n.SCRIPT_TAG].close,do:this[n.SCRIPT_TAG].do},this[n.SCRIPT_TAG|n.COMMENT|n.MULTILINE]={do:function(){a("*/")&&(t=n.SCRIPT_TAG)}},this[n.SCRIPT_TAG|n.COMMENT|n.ONELINE]={do:function(){a(`
`)&&(t=n.SCRIPT_TAG)}},this[n.SCRIPT_TAG|n.SINGLE_QUOTE]={do:function(){a("'")&&(t=n.SCRIPT_TAG)}},this[n.SCRIPT_TAG|n.DOUBLE_QUOTE]={do:function(){a('"')&&(t=n.SCRIPT_TAG)}},this[n.SCRIPT_TAG|n.MULTILINE]={do:function(){a("`")&&(t=n.SCRIPT_TAG)}},this[n.COMMENT]={enter:function(){u.set(r)},exit:function(){var p=f(u.get(),r-2);!p||l.onComment(p)},do:function(){a("-->",f(r-2,r+1))&&(t=n.TEXT)}}},T={},A;r<e.length;r++,o=e[r])A=!1,T=m[t],T&&(s!=t&&T.enter&&T.enter(),s=t,T.do&&(T.do(),t!=s&&T.exit&&(T.exit(),A=!0)));!A&&T.exit&&T.exit()}function C(e){(e.type==="root"||e.type==="node")&&(delete e.parent,delete e.last,e.children&&e.children.forEach(function(l){C(l)}))}function X(e){let l={"!DOCTYPE":!0,"!doctype":!0,area:!0,base:!0,br:!0,col:!0,command:!0,embed:!0,hr:!0,img:!0,input:!0,keygen:!0,link:!0,meta:!0,param:!0,source:!0,track:!0,wbr:!0,circle:!0,ellipse:!0,line:!0,path:!0,polygon:!0,polyline:!0,rect:!0,stop:!0,use:!0},r={type:"root"},o=r,t={};return t.onOpenTag=function(s){var c={type:"node",tagName:s};c.parent=o,o.children||(o.children=[]),o.children.push(c),o=c},t.onOpenTagExit=function(){l[o.tagName]&&(o=o.parent)},t.onCloseTag=function(){o=o.parent},t.onSelfClose=function(){o=o.parent},t.onAttrName=function(s){o.attrs||(o.attrs={}),o.last=s,o.attrs[s]=""},t.onAttrValue=function(s){o.attrs[o.last]=s},t.onText=function(s){o.children||(o.children=[]);var c={type:"text",text:V(s)};o.children.push(c)},t.onComment=function(s){o.children||(o.children=[]);var c={type:"comment",value:s};o.children.push(c)},j(e,t),C(r),r}var H="/public_assets/images/fig1.jpg",W="/public_assets/images/fig2.jpg",$="/public_assets/images/fig3.jpg",q="/public_assets/images/fig4.jpg";const F=[H,W,$,q],g=e=>e.includes("cm")?parseInt(e.replace("cm"))*566.9291338583:e.includes("pt")?parseInt(e.replace("pt"))*20:0,J=[{id:"alignment",key:"text-align",isStatic:!0,values:{center:i.exports.AlignmentType.CENTER,left:i.exports.AlignmentType.LEFT,right:i.exports.AlignmentType.RIGHT,justify:i.exports.AlignmentType.JUSTIFIED}},{id:"color",key:"color",isStatic:!1,value:e=>i.exports.hexColorValue(e)},{id:"font",key:"font-family",isStatic:!1,value:e=>e},{id:"size",key:"font-size",isStatic:!1,value:e=>parseFloat(e.replace("pt",""))},{id:"indent",key:"padding",isStatic:!1,value:e=>{const l=e.split(" ");return{top:g(l[0]),right:g(l[1]),bottom:g(l[2]),left:g(l[3])}}},{id:"highlight",key:"background-color",isStatic:!1,value:e=>i.exports.hexColorValue(e)},{id:"width",key:"width",isStatic:!1,value:e=>({size:g(e),type:i.exports.WidthType.DXA})},{id:"spacing",key:"line-height",isStatic:!1,value:e=>({line:typeof e=="string"?240:parseFloat((e*240/100).toFixed(2))})}],K=(e,l)=>{const r=J.find(o=>o.key===e);return r?[r.id,r.isStatic&&r.values[l]?r.values[l]:r.value(l)]:null},Z=e=>{let l=e.split(";"),r=l.length,o={style:{}},t,s,c;for(;r--;)if(t=l[r].split(":"),s=N.exports.trim(t[0]),c=N.exports.trim(t[1]),s.length>0&&c.length>0){const a=K(s,c);a!==null?o.style[a[0]]=a[1]:o.style[s]=c}return o.style},ee=e=>{let{state:l}=e;(function r(o){let t=o.findIndex(c=>c.key==="span");if(t===-1)return;let s=o.slice(t);e.styles||(e.styles={current:{}}),Object.assign(e.styles.current,s[0].data),r(s.slice(1))})(l)},G=(e,l={},r=[],o=0)=>{let t={type:e.type==="node"?e.tagName:e.type,styles:{current:{}},level:o},s={};e.attrs&&(t.attributes=e.attrs,t.attributes.style&&(t.styles.current=Z(t.attributes.style),s=t.styles.current));let c={};c.key=t.type,c.data=s;const a=[...r,c];switch(e.children&&(t.children=e.children.map(f=>G(f,s,a,o+1))),t.type){case"text":t.state=a,ee(t),t.value=e.text;break;case"strong":t.children[0].isStrong=!0,t=t.children[0],Object.assign(t.styles.current,l);break;case"span":let f=[],E=0;t.children=t.children.filter((d,u)=>d.children?(d.children.forEach(m=>f.push({i:u,c:m})),!1):!0),f.length>0&&f.forEach(d=>{t.children.splice(d.i+E,0,d.c),E+=1});break;case"sub":t.children[0].subScript=!0;break;case"u":t.children[0].underline=!0;break;case"s":t.children[0].strike=!0;break;default:t.__v=1}return t.state=a,t},b=(e,l={p:0})=>{let r=null;if(e.children&&(r=e.children.map(o=>b(o,l)).filter(o=>o!==null)),e.type==="root")return{section:r};if(e.type==="p"){const o=Object.entries(e.styles)[0];let t={};o!==null&&(t=o[1]);let s=[];return r.forEach(c=>{Array.isArray(c)?s=[...s,...c]:s.push(c)}),e.level===1&&(l.p+=1,l.p!==1&&s.unshift(new i.exports.TextRun({text:`[${re(l.p-1,4)}]  `})),l.p===1&&(t.heading=i.exports.HeadingLevel.HEADING_1)),"size"in t&&(t.size*=2),new i.exports.Paragraph(y({children:s},t))}else if(e.type==="text"){const o=e;if(o.value===" ")return null;const t=Object.entries(e.styles)[0];let s={};return t!==null&&(s=t[1]),s||(s={}),s.bold=!!o.isStrong,s.subScript=!!o.subScript,o.underline&&(s.underline={}),o.strike&&(s.strike={}),"size"in s&&(s.size*=2),new i.exports.TextRun(y({text:o.value},s))}else{if(e.type==="span")return r;if(e.type==="strong"||e.type==="mfenced")return r[0];if(e.type==="sub")return r[0];if(e.type==="u")return r[0];if(e.type==="s")return r[0];if(e.type==="img")return new i.exports.TextRun({text:"(---Image here---)"});if(e.type==="comment")return new i.exports.PageBreak;if(e.type==="table")return r[0];if(e.type==="div")return r;if(e.type==="br")return new i.exports.TextRun({text:"break",break:1});if(e.type==="tbody")return new i.exports.Table({rows:r});if(e.type==="tr")return new i.exports.TableRow({children:r});if(e.type==="td"){const o={};return Object.entries(e.attributes).forEach(([t,s])=>{t==="colspan"&&(o.columnSpan=s)}),new i.exports.TableCell(y(y({children:r},o),e.styles.current))}else{if(e.type==="math")return new i.exports.Math({children:r});if(e.type==="mi"||e.type==="mo"||e.type==="mn")return new i.exports.MathRun(e.children[0].value);if(e.type==="msqrt")return new i.exports.MathRadical({children:r});if(e.type==="msup")return new i.exports.MathSuperScript({children:[r[0]],superScript:[r[1]]});if(e.type==="msub")return new i.exports.MathSubScript({children:[r[0]],subScript:[r[1]]});if(e.type==="mrow")return r;if(e.type==="mfrac")return new i.exports.MathFraction({numerator:r[0],denominator:r[1]})}}return null},te=e=>new Promise(l=>{var r=new XMLHttpRequest;r.open("GET",e,!0),r.responseType="blob",r.onload=function(){console.log(this.response);var o=new FileReader;o.onload=function(s){var c=s.target.result;l(c)};var t=this.response;o.readAsDataURL(t)},r.send()});function re(e,l){for(e=e.toString();e.length<l;)e="0"+e;return e}const ne=async e=>{const l=G(e);let r=b(l).section.filter(s=>s);F&&(await Promise.all(Object.entries(F).map(async c=>{const a={};return a[c[0]]=await te(c[1]),a}))).forEach(c=>{const a=Object.entries(c)[0],f=new i.exports.Paragraph({children:[new i.exports.TextRun({text:a[0]})],alignment:i.exports.AlignmentType.CENTER});r.push(f);const E=new i.exports.Paragraph({alignment:i.exports.AlignmentType.CENTER,children:[new i.exports.ImageRun({data:a[1],transformation:{width:500,height:500}}),new i.exports.PageBreak],spacing:{before:200}});r.push(E)});const o=new i.exports.Document({sections:[{properties:{lineNumbers:{countBy:5,restart:i.exports.LineNumberRestartFormat.NEW_PAGE,size:10},page:{margin:{top:1500,right:1500,bottom:1500,left:1500},pageNumbers:{start:1,formatType:i.exports.NumberFormat.DECIMAL},size:{height:16839}}},headers:{default:new i.exports.Header({children:[new i.exports.Table({alignment:i.exports.AlignmentType.CENTER,width:{size:100,type:i.exports.WidthType.PERCENTAGE},rows:[new i.exports.TableRow({children:[new i.exports.TableCell({columnSpan:1,children:[new i.exports.Paragraph({children:[new i.exports.TextRun({text:"IPACTORY",size:24})],alignment:i.exports.AlignmentType.LEFT})],width:{size:33,type:i.exports.WidthType.PERCENTAGE},borders:{top:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"}}}),new i.exports.TableCell({columnSpan:1,children:[new i.exports.Paragraph({size:12,font:"Times New Roman",alignment:i.exports.AlignmentType.CENTER,children:[new i.exports.TextRun({children:[new i.exports.TextRun({text:"- ",size:24,font:"Times New Roman"}),new i.exports.TextRun({children:[i.exports.PageNumber.CURRENT],font:"Times New Roman",size:24}),new i.exports.TextRun({text:" -",size:24,font:"Times New Roman"})]})]})],verticalAlign:i.exports.VerticalAlign.CENTER,borders:{top:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"}},width:{size:33,type:i.exports.WidthType.PERCENTAGE}}),new i.exports.TableCell({columnSpan:1,children:[new i.exports.Paragraph({alignment:i.exports.AlignmentType.RIGHT,children:[new i.exports.TextRun({text:"259889523",size:24})]})],verticalAlign:i.exports.VerticalAlign.CENTER,borders:{top:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:i.exports.BorderStyle.NONE,size:0,color:"FFFFFF"}},width:{size:33,type:i.exports.WidthType.PERCENTAGE}})]})]})]})},children:[...r]}]}),t="application/vnd.openxmlformats-officedocument.wordprocessingml.document";return new Promise((s,c)=>i.exports.Packer.toBlob(o).then(a=>{const f=a.slice(0,a.size,t);M.exports.saveAs(f,"output.docx"),s(f)}).catch(a=>{console.error(a),c(a)}))};var R=(e,l)=>{const r=e.__vccOpts||e;for(const[o,t]of l)r[o]=t;return r};const ie={name:"TestComponent",data(){return{file:null,text:null,document:null,parsedText:null}},methods:{async submitForm(){this.document=await ne(this.parsedText)},fileUploaded(e){const[l]=e.target.files;this.file=l;const r=new FileReader;r.onload=(()=>o=>{this.text=o.target.result.replace(/\r?\n|\r/g,"").replace(/\s\s+/g," ").replace(/>\s+</g,"><")})(),r.readAsText(l)}},watch:{text(e){e!==null&&(this.parsedText=X(e.replace(/\r?\n|\r/g,"").replace(/\s\s+/g," ").replace(/>\s+</g,"><")),this.submitForm())}}},w=e=>(k("data-v-372eca44"),e=e(),z(),e),se={class:"input"},oe=w(()=>_("label",{for:"file"},"Upload an html file",-1)),le=w(()=>_("div",{id:"container"},null,-1));function ae(e,l,r,o,t,s){return O(),L("div",se,[_("form",{onSubmit:l[1]||(l[1]=U((...c)=>s.submitForm&&s.submitForm(...c),["prevent"]))},[oe,_("input",{type:"file",id:"file",onChange:l[0]||(l[0]=(...c)=>s.fileUploaded&&s.fileUploaded(...c)),required:""},null,32)],32),le])}var ce=R(ie,[["render",ae],["__scopeId","data-v-372eca44"]]);const pe={name:"AppComponent",components:{Test:ce}};function ue(e,l,r,o,t,s){const c=B("Test");return O(),L("main",null,[Y(c)])}var fe=R(pe,[["render",ue]]);D(fe).mount("#app");