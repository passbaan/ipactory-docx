var b=Object.defineProperty;var S=Object.getOwnPropertySymbols;var R=Object.prototype.hasOwnProperty,P=Object.prototype.propertyIsEnumerable;var y=(r,s,t)=>s in r?b(r,s,{enumerable:!0,configurable:!0,writable:!0,value:t}):r[s]=t,h=(r,s)=>{for(var t in s||(s={}))R.call(s,t)&&y(r,t,s[t]);if(S)for(var t of S(s))P.call(s,t)&&y(r,t,s[t]);return r};import{b as a,l as L,d as M,o as G,c as O,a as m,w as U,r as k,e as w,f as Y}from"./vendor.591d3f58.js";const F=function(){const s=document.createElement("link").relList;if(s&&s.supports&&s.supports("modulepreload"))return;for(const n of document.querySelectorAll('link[rel="modulepreload"]'))i(n);new MutationObserver(n=>{for(const l of n)if(l.type==="childList")for(const u of l.addedNodes)u.tagName==="LINK"&&u.rel==="modulepreload"&&i(u)}).observe(document,{childList:!0,subtree:!0});function t(n){const l={};return n.integrity&&(l.integrity=n.integrity),n.referrerpolicy&&(l.referrerPolicy=n.referrerpolicy),n.crossorigin==="use-credentials"?l.credentials="include":n.crossorigin==="anonymous"?l.credentials="omit":l.credentials="same-origin",l}function i(n){if(n.ep)return;n.ep=!0;const l=t(n);fetch(n.href,l)}};F();var e={NONE:0,OPEN_TAG:1<<1,CLOSE_TAG:1<<2,SELF_CLOSING:1<<3,ATTR_NAME:1<<4,ATTR_VALUE:1<<5,TEXT:1<<6,COMMENT:1<<7,ONELINE:1<<8,MULTILINE:1<<9,SINGLE_QUOTE:1<<10,DOUBLE_QUOTE:1<<11,STYLE_TAG:1<<12,SCRIPT_TAG:1<<13};const Q=function(){var r=document.createElement("div");function s(t){return t&&typeof t=="string"&&(t=t.replace(/<script[^>]*>([\S\s]*?)<\/script>/gim,""),t=t.replace(/<\/?\w(?:[^"'>]|"[^"]*"|'[^']*')*>/gim,""),r.innerHTML=t,t=r.textContent,r.textContent=""),t}return s}();function B(r,s){var t=0,i=r[t],n=e.TEXT,l=e.NONE,u=e.NONE;function o(c,_){return c.length<=1&&!_?c===i:c===(_||r.substring(t-c.length,t))}function T(c,_){return r.substring(c,_)}function d(c){return c===n}function E(c){return c===l}var f=[0];f.set=function(c){f[f.length-1]=c},f.get=function(){return f[f.length-1]};for(var v=new function(){this[e.TEXT]={enter:function(){E(e.OPEN_TAG|e.CLOSE_TAG)||E(e.STYLE_TAG)||E(e.SCRIPT_TAG)||f.set(t)},exit:function(){var c=T(f.get(),t);!c||s.onText(c)},do:function(){o("<")&&(n=e.OPEN_TAG|e.CLOSE_TAG)}},this[e.OPEN_TAG|e.CLOSE_TAG]={enter:function(){f.push(t)},exit:function(){var c=f.pop();d(e.OPEN_TAG)&&f.set(c)},do:function(){o("!--")?n=e.COMMENT:o("/")?n=e.CLOSE_TAG:o(" ")?n=e.TEXT:!o("!")&&!o("-")&&(n=e.OPEN_TAG)}},this[e.OPEN_TAG]={exit:function(){var c=T(f.get(),t);!c||(c==="style"&&(u=e.STYLE_TAG),c==="script"&&(u=e.SCRIPT_TAG),s.onOpenTag(c))},do:function(){o(">")?n=e.SELF_CLOSING:o(" ")&&(n=e.ATTR_NAME)}},this[e.CLOSE_TAG]={enter:function(){f.set(t)},exit:function(){var c=T(f.get(),t);!c||s.onCloseTag(c)},do:function(){o(">")&&(n=e.TEXT)}},this[e.SELF_CLOSING]={enter:function(){f.push(t)},exit:function(){var c=f.pop();t=c-1,u&e.STYLE_TAG&&(n=e.OPEN_TAG|e.STYLE_TAG),u&e.SCRIPT_TAG&&(n=e.OPEN_TAG|e.SCRIPT_TAG)},do:function(){o("/>",T(t-2,t))?this.mark():this.close(),n=e.TEXT},mark:function(){s.onSelfClose()},close:function(){s.onOpenTagExit()}},this[e.ATTR_NAME]={enter:function(){f.set(t)},exit:function(){if(!o("/",T(t-1,t))){var c=T(f.get(),t);!c||s.onAttrName(c)}},do:function(){o(" ")||o("/")?this.repeat():o(">")?n=e.SELF_CLOSING:o("=")&&(n=e.ATTR_VALUE)},repeat:function(){this.exit(),f.set(t+1)}},this[e.ATTR_VALUE]={enter:function(){o("'")?(n|=e.SINGLE_QUOTE,f.set(t+1)):o('"')?(n|=e.DOUBLE_QUOTE,f.set(t+1)):f.set(t)},exit:function(){var c=T(f.get(t),t);!c||s.onAttrValue(c)},do:function(){o(" ")?n=e.ATTR_NAME:o(">")&&(n=e.SELF_CLOSING)}},this[e.ATTR_VALUE|e.SINGLE_QUOTE]={exit:this[e.ATTR_VALUE].exit,do:function(){o("'")&&(n=e.ATTR_NAME)}},this[e.ATTR_VALUE|e.DOUBLE_QUOTE]={exit:this[e.ATTR_VALUE].exit,do:function(){o('"')&&(n=e.ATTR_NAME)}},this[e.STYLE_TAG]={do:function(){o("/*")?n=e.STYLE_TAG|e.COMMENT:o("'")?n=e.STYLE_TAG|e.SINGLE_QUOTE:o('"')?n=e.STYLE_TAG|e.DOUBLE_QUOTE:o("<")&&o("</style>",T(t,t+8))&&(n=e.TEXT,t--,this.close())},close:function(){var c=f.pop();f.set(c),u^=e.STYLE_TAG}},this[e.OPEN_TAG|e.STYLE_TAG]={enter:function(){f.push(t)},do:this[e.STYLE_TAG].do},this[e.STYLE_TAG|e.COMMENT]={do:function(){o("*/")&&(n=e.STYLE_TAG)}},this[e.STYLE_TAG|e.SINGLE_QUOTE]={do:function(){o("'")&&(n=e.STYLE_TAG)}},this[e.STYLE_TAG|e.DOUBLE_QUOTE]={do:function(){o('"')&&(n=e.STYLE_TAG)}},this[e.SCRIPT_TAG]={do:function(){o("/*")?n=e.SCRIPT_TAG|e.COMMENT|e.MULTILINE:o("//")?n=e.SCRIPT_TAG|e.COMMENT|e.ONELINE:o("'")?n=e.SCRIPT_TAG|e.SINGLE_QUOTE:o('"')?n=e.SCRIPT_TAG|e.DOUBLE_QUOTE:o("`")?n=e.SCRIPT_TAG|e.MULTILINE:o("<")&&o("<\/script>",T(t,t+9))&&(n=e.TEXT,t--,this.close())},close:function(){var c=f.pop();f.set(c),u^=e.SCRIPT_TAG}},this[e.OPEN_TAG|e.SCRIPT_TAG]={enter:function(){f.push(t)},exit:this[e.SCRIPT_TAG].close,do:this[e.SCRIPT_TAG].do},this[e.SCRIPT_TAG|e.COMMENT|e.MULTILINE]={do:function(){o("*/")&&(n=e.SCRIPT_TAG)}},this[e.SCRIPT_TAG|e.COMMENT|e.ONELINE]={do:function(){o(`
`)&&(n=e.SCRIPT_TAG)}},this[e.SCRIPT_TAG|e.SINGLE_QUOTE]={do:function(){o("'")&&(n=e.SCRIPT_TAG)}},this[e.SCRIPT_TAG|e.DOUBLE_QUOTE]={do:function(){o('"')&&(n=e.SCRIPT_TAG)}},this[e.SCRIPT_TAG|e.MULTILINE]={do:function(){o("`")&&(n=e.SCRIPT_TAG)}},this[e.COMMENT]={enter:function(){f.set(t)},exit:function(){var c=T(f.get(),t-2);!c||s.onComment(c)},do:function(){o("-->",T(t-2,t+1))&&(n=e.TEXT)}}},p={},g;t<r.length;t++,i=r[t])g=!1,p=v[n],p&&(l!=n&&p.enter&&p.enter(),l=n,p.do&&(p.do(),n!=l&&p.exit&&(p.exit(),g=!0)));!g&&p.exit&&p.exit()}function N(r){(r.type==="root"||r.type==="node")&&(delete r.parent,delete r.last,r.children&&r.children.forEach(function(s){N(s)}))}function V(r){let s={"!DOCTYPE":!0,"!doctype":!0,area:!0,base:!0,br:!0,col:!0,command:!0,embed:!0,hr:!0,img:!0,input:!0,keygen:!0,link:!0,meta:!0,param:!0,source:!0,track:!0,wbr:!0,circle:!0,ellipse:!0,line:!0,path:!0,polygon:!0,polyline:!0,rect:!0,stop:!0,use:!0},t={type:"root"},i=t,n={};return n.onOpenTag=function(l){var u={type:"node",tagName:l};u.parent=i,i.children||(i.children=[]),i.children.push(u),i=u},n.onOpenTagExit=function(){s[i.tagName]&&(i=i.parent)},n.onCloseTag=function(){i=i.parent},n.onSelfClose=function(){i=i.parent},n.onAttrName=function(l){i.attrs||(i.attrs={}),i.last=l,i.attrs[l]=""},n.onAttrValue=function(l){i.attrs[i.last]=l},n.onText=function(l){i.children||(i.children=[]);var u={type:"text",text:Q(l)};i.children.push(u)},n.onComment=function(l){i.children||(i.children=[]);var u={type:"comment",value:l};i.children.push(u)},B(r,n),N(t),t}const A=r=>r.includes("cm")?parseInt(r.replace("cm"))*566.9291338583:r.includes("pt")?parseInt(r.replace("pt"))*20:0,D=[{id:"alignment",key:"text-align",isStatic:!0,values:{center:a.exports.AlignmentType.CENTER,left:a.exports.AlignmentType.LEFT,right:a.exports.AlignmentType.RIGHT,justify:a.exports.AlignmentType.JUSTIFIED}},{id:"color",key:"color",isStatic:!1,value:r=>a.exports.hexColorValue(r)},{id:"font",key:"font-family",isStatic:!1,value:r=>r},{id:"size",key:"font-size",isStatic:!1,value:r=>parseFloat(r.replace("pt",""))},{id:"indent",key:"padding",isStatic:!1,value:r=>{const s=r.split(" ");return{top:A(s[0]),right:A(s[1]),bottom:A(s[2]),left:A(s[3])}}},{id:"highlight",key:"background-color",isStatic:!1,value:r=>a.exports.hexColorValue(r)}],j=(r,s)=>{const t=D.find(i=>i.key===r);return t?[t.id,t.isStatic&&t.values[s]?t.values[s]:t.value(s)]:null},X=r=>{let s=r.split(";"),t=s.length,i={style:{}},n,l,u;for(;t--;)if(n=s[t].split(":"),l=L.exports.trim(n[0]),u=L.exports.trim(n[1]),l.length>0&&u.length>0){const o=j(l,u);o!==null?i.style[o[0]]=o[1]:i.style[l]=u}return i.style},C=(r,s={},t=[])=>{let i={type:r.type==="node"?r.tagName:r.type,styles:{current:null}},n={};r.attrs&&(i.attributes=r.attrs,i.attributes.style&&(i.styles.current=X(i.attributes.style),n=i.styles.current));let l={};l[i.type]=n;const u=[...t,l];switch(r.children&&(i.children=r.children.map(o=>C(o,n,u))),i.type){case"text":i.value=r.text,i.styles.current=s;break;case"strong":i.children[0].isStrong=!0,i=i.children[0],Object.assign(i.styles.current,s);break;case"span":i.styles.current=s;let o=[],T=0;i.children=i.children.filter((d,E)=>d.children?(console.log("file: doc.service.js | line 148 | text.children=text.children.filter | ch",d),d.children.forEach(f=>o.push({i:E,c:f})),!1):!0),o.length>0&&o.forEach(d=>{i.children.splice(d.i+T,0,d.c),T+=1});break;case"sub":i.children[0].subScript=!0;break;case"u":i.styles.current=s,i.children[0].underline=!0;break;case"s":i.children[0].strike=!0;break;default:i.__v=1}return i.state=u,i},x=r=>{let s=null;if(r.children&&(s=r.children.map(t=>x(t)).filter(t=>t!==null)),r.type==="root")return{section:s};if(r.type==="p"){console.log("file: doc.service.js | line 190 | generate | x",r);const t=Object.entries(r.styles)[0];let i={};t!==null&&(i=t[1]),console.log("file: doc.service.js | line 199 | generate | children",s);let n=[];return s.forEach(l=>{Array.isArray(l)?n=[...n,...l]:n.push(l)}),new a.exports.Paragraph(h({children:n},i))}else if(r.type==="text"){const t=r;if(t.value===" ")return null;const i=Object.entries(r.styles)[0];let n={};return i!==null&&(n=i[1]),n||(n={}),n.bold=!!t.isStrong,n.subScript=!!t.subScript,t.underline&&(n.underline={}),t.strike&&(n.strike={}),new a.exports.TextRun(h({text:t.value},n))}else{if(r.type==="span")return s;if(r.type==="strong")return s[0];if(r.type==="sub")return s[0];if(r.type==="u")return s[0];if(r.type==="s")return s[0];if(r.type==="img")return new a.exports.TextRun({text:"(---Image here---)"});if(r.type==="comment")return new a.exports.TextRun({text:"COMMENT HERE"});if(r.type==="table")return s[0];if(r.type==="div")return s;if(r.type==="br")return new a.exports.TextRun({text:"break",break:1});if(r.type==="tbody")return new a.exports.Table({rows:s});if(r.type==="tr")return new a.exports.TableRow({children:s});if(r.type==="td"){const t={};return Object.entries(r.attributes).forEach(([i,n])=>{i==="colspan"&&(t.columnSpan=n)}),new a.exports.TableCell(h({children:s},t))}}return console.log("here",r),null},$=r=>{const s=C(r);console.log("file: doc.service.js | line 186 | createNew | x",s);let t=x(s).section.filter(l=>l);const i=new a.exports.Document({sections:[{properties:{page:{pageNumbers:{start:1,formatType:a.exports.NumberFormat.DECIMAL},margin:{top:0,right:200,bottom:0,left:200}}},children:t}]}),n="application/vnd.openxmlformats-officedocument.wordprocessingml.document";return new Promise((l,u)=>a.exports.Packer.toBlob(i).then(o=>{const T=o.slice(0,o.size,n);l(T)}).catch(o=>{console.error(o),u(o)}))};var I=(r,s)=>{const t=r.__vccOpts||r;for(const[i,n]of s)t[i]=n;return t};const H={name:"TestComponent",data(){return{file:null,text:null,document:null,parsedText:null}},methods:{async submitForm(){this.document=await $(this.parsedText),M.exports.renderAsync(this.document,document.getElementById("container")).then(()=>console.log("docx: finished"))},fileUploaded(r){const[s]=r.target.files;this.file=s;const t=new FileReader;t.onload=(()=>i=>{this.text=i.target.result.replace(/\r?\n|\r/g,"").replace(/\s\s+/g," ").replace(/>\s+</g,"><")})(),t.readAsText(s)}},watch:{text(r){r!==null&&(this.parsedText=V(r.replace(/\r?\n|\r/g,"").replace(/\s\s+/g," ").replace(/>\s+</g,"><")))}}},q=m("button",{type:"submit"},"Start",-1),z=m("div",{id:"container"},null,-1);function J(r,s,t,i,n,l){return G(),O("div",null,[m("form",{onSubmit:s[1]||(s[1]=U((...u)=>l.submitForm&&l.submitForm(...u),["prevent"]))},[m("input",{type:"file",onChange:s[0]||(s[0]=(...u)=>l.fileUploaded&&l.fileUploaded(...u)),required:""},null,32),q],32),z])}var K=I(H,[["render",J]]);const W={name:"AppComponent",components:{Test:K}};function Z(r,s,t,i,n,l){const u=k("Test");return G(),O("main",null,[w(u)])}var ee=I(W,[["render",Z]]);Y(ee).mount("#app");
