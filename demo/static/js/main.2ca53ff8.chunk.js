/*! For license information please see main.2ca53ff8.chunk.js.LICENSE.txt */
(this["webpackJsonpimec.ai"]=this["webpackJsonpimec.ai"]||[]).push([[0],{117:function(e,t,n){},127:function(e,t){},154:function(e,t,n){"use strict";n.r(t);var r=n(0),a=n.n(r),o=n(9),i=n.n(o),c=(n(87),n(12)),s=n(14),u=n.n(s),l=n(181),p=n(182),f=n(183),h=n(55),d=n(71),m=n.n(d);function g(){g=function(){return e};var e={},t=Object.prototype,n=t.hasOwnProperty,r="function"==typeof Symbol?Symbol:{},a=r.iterator||"@@iterator",o=r.asyncIterator||"@@asyncIterator",i=r.toStringTag||"@@toStringTag";function c(e,t,n){return Object.defineProperty(e,t,{value:n,enumerable:!0,configurable:!0,writable:!0}),e[t]}try{c({},"")}catch(k){c=function(e,t,n){return e[t]=n}}function s(e,t,n,r){var a=t&&t.prototype instanceof p?t:p,o=Object.create(a.prototype),i=new j(r||[]);return o._invoke=function(e,t,n){var r="suspendedStart";return function(a,o){if("executing"===r)throw new Error("Generator is already running");if("completed"===r){if("throw"===a)throw o;return T()}for(n.method=a,n.arg=o;;){var i=n.delegate;if(i){var c=b(i,n);if(c){if(c===l)continue;return c}}if("next"===n.method)n.sent=n._sent=n.arg;else if("throw"===n.method){if("suspendedStart"===r)throw r="completed",n.arg;n.dispatchException(n.arg)}else"return"===n.method&&n.abrupt("return",n.arg);r="executing";var s=u(e,t,n);if("normal"===s.type){if(r=n.done?"completed":"suspendedYield",s.arg===l)continue;return{value:s.arg,done:n.done}}"throw"===s.type&&(r="completed",n.method="throw",n.arg=s.arg)}}}(e,n,i),o}function u(e,t,n){try{return{type:"normal",arg:e.call(t,n)}}catch(k){return{type:"throw",arg:k}}}e.wrap=s;var l={};function p(){}function f(){}function h(){}var d={};c(d,a,(function(){return this}));var m=Object.getPrototypeOf,v=m&&m(m(O([])));v&&v!==t&&n.call(v,a)&&(d=v);var y=h.prototype=p.prototype=Object.create(d);function x(e){["next","throw","return"].forEach((function(t){c(e,t,(function(e){return this._invoke(t,e)}))}))}function w(e,t){var r;this._invoke=function(a,o){function i(){return new t((function(r,i){!function r(a,o,i,c){var s=u(e[a],e,o);if("throw"!==s.type){var l=s.arg,p=l.value;return p&&"object"==typeof p&&n.call(p,"__await")?t.resolve(p.__await).then((function(e){r("next",e,i,c)}),(function(e){r("throw",e,i,c)})):t.resolve(p).then((function(e){l.value=e,i(l)}),(function(e){return r("throw",e,i,c)}))}c(s.arg)}(a,o,r,i)}))}return r=r?r.then(i,i):i()}}function b(e,t){var n=e.iterator[t.method];if(void 0===n){if(t.delegate=null,"throw"===t.method){if(e.iterator.return&&(t.method="return",t.arg=void 0,b(e,t),"throw"===t.method))return l;t.method="throw",t.arg=new TypeError("The iterator does not provide a 'throw' method")}return l}var r=u(n,e.iterator,t.arg);if("throw"===r.type)return t.method="throw",t.arg=r.arg,t.delegate=null,l;var a=r.arg;return a?a.done?(t[e.resultName]=a.value,t.next=e.nextLoc,"return"!==t.method&&(t.method="next",t.arg=void 0),t.delegate=null,l):a:(t.method="throw",t.arg=new TypeError("iterator result is not an object"),t.delegate=null,l)}function E(e){var t={tryLoc:e[0]};1 in e&&(t.catchLoc=e[1]),2 in e&&(t.finallyLoc=e[2],t.afterLoc=e[3]),this.tryEntries.push(t)}function L(e){var t=e.completion||{};t.type="normal",delete t.arg,e.completion=t}function j(e){this.tryEntries=[{tryLoc:"root"}],e.forEach(E,this),this.reset(!0)}function O(e){if(e){var t=e[a];if(t)return t.call(e);if("function"==typeof e.next)return e;if(!isNaN(e.length)){var r=-1,o=function t(){for(;++r<e.length;)if(n.call(e,r))return t.value=e[r],t.done=!1,t;return t.value=void 0,t.done=!0,t};return o.next=o}}return{next:T}}function T(){return{value:void 0,done:!0}}return f.prototype=h,c(y,"constructor",h),c(h,"constructor",f),f.displayName=c(h,i,"GeneratorFunction"),e.isGeneratorFunction=function(e){var t="function"==typeof e&&e.constructor;return!!t&&(t===f||"GeneratorFunction"===(t.displayName||t.name))},e.mark=function(e){return Object.setPrototypeOf?Object.setPrototypeOf(e,h):(e.__proto__=h,c(e,i,"GeneratorFunction")),e.prototype=Object.create(y),e},e.awrap=function(e){return{__await:e}},x(w.prototype),c(w.prototype,o,(function(){return this})),e.AsyncIterator=w,e.async=function(t,n,r,a,o){void 0===o&&(o=Promise);var i=new w(s(t,n,r,a),o);return e.isGeneratorFunction(n)?i:i.next().then((function(e){return e.done?e.value:i.next()}))},x(y),c(y,i,"Generator"),c(y,a,(function(){return this})),c(y,"toString",(function(){return"[object Generator]"})),e.keys=function(e){var t=[];for(var n in e)t.push(n);return t.reverse(),function n(){for(;t.length;){var r=t.pop();if(r in e)return n.value=r,n.done=!1,n}return n.done=!0,n}},e.values=O,j.prototype={constructor:j,reset:function(e){if(this.prev=0,this.next=0,this.sent=this._sent=void 0,this.done=!1,this.delegate=null,this.method="next",this.arg=void 0,this.tryEntries.forEach(L),!e)for(var t in this)"t"===t.charAt(0)&&n.call(this,t)&&!isNaN(+t.slice(1))&&(this[t]=void 0)},stop:function(){this.done=!0;var e=this.tryEntries[0].completion;if("throw"===e.type)throw e.arg;return this.rval},dispatchException:function(e){if(this.done)throw e;var t=this;function r(n,r){return i.type="throw",i.arg=e,t.next=n,r&&(t.method="next",t.arg=void 0),!!r}for(var a=this.tryEntries.length-1;a>=0;--a){var o=this.tryEntries[a],i=o.completion;if("root"===o.tryLoc)return r("end");if(o.tryLoc<=this.prev){var c=n.call(o,"catchLoc"),s=n.call(o,"finallyLoc");if(c&&s){if(this.prev<o.catchLoc)return r(o.catchLoc,!0);if(this.prev<o.finallyLoc)return r(o.finallyLoc)}else if(c){if(this.prev<o.catchLoc)return r(o.catchLoc,!0)}else{if(!s)throw new Error("try statement without catch or finally");if(this.prev<o.finallyLoc)return r(o.finallyLoc)}}}},abrupt:function(e,t){for(var r=this.tryEntries.length-1;r>=0;--r){var a=this.tryEntries[r];if(a.tryLoc<=this.prev&&n.call(a,"finallyLoc")&&this.prev<a.finallyLoc){var o=a;break}}o&&("break"===e||"continue"===e)&&o.tryLoc<=t&&t<=o.finallyLoc&&(o=null);var i=o?o.completion:{};return i.type=e,i.arg=t,o?(this.method="next",this.next=o.finallyLoc,l):this.complete(i)},complete:function(e,t){if("throw"===e.type)throw e.arg;return"break"===e.type||"continue"===e.type?this.next=e.arg:"return"===e.type?(this.rval=this.arg=e.arg,this.method="return",this.next="end"):"normal"===e.type&&t&&(this.next=t),l},finish:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.finallyLoc===e)return this.complete(n.completion,n.afterLoc),L(n),l}},catch:function(e){for(var t=this.tryEntries.length-1;t>=0;--t){var n=this.tryEntries[t];if(n.tryLoc===e){var r=n.completion;if("throw"===r.type){var a=r.arg;L(n)}return a}}throw new Error("illegal catch attempt")},delegateYield:function(e,t,n){return this.delegate={iterator:O(e),resultName:t,nextLoc:n},"next"===this.method&&(this.arg=void 0),l}},e}function v(){return(v=Object(h.a)(g().mark((function e(t,n,r){var a,o;return g().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return a=new Date,e.t0=Promise,e.next=4,fetch(t+"/dalle",{method:"POST",headers:{"Bypass-Tunnel-Reminder":"go",mode:"no-cors"},body:JSON.stringify({text:n,num_images:r})}).then((function(e){if(!e.ok)throw Error(e.statusText);return e}));case 4:return e.t1=e.sent.text(),e.t2=new Promise((function(e,t){return setTimeout((function(){return t(new Error("Timeout"))}),6e4)})),e.t3=[e.t1,e.t2],e.next=9,e.t0.race.call(e.t0,e.t3);case 9:return o=e.sent,e.abrupt("return",{executionTime:Math.round(100*((new Date-a)/1e3+Number.EPSILON))/100,generatedImgs:m.a.parse(o)});case 11:case"end":return e.stop()}}),e)})))).apply(this,arguments)}function y(){return(y=Object(h.a)(g().mark((function e(t){return g().wrap((function(e){for(;;)switch(e.prev=e.next){case 0:return e.next=2,fetch(t,{headers:{"Bypass-Tunnel-Reminder":"go",mode:"no-cors"}}).then((function(e){return!0})).catch((function(){return!1}));case 2:return e.abrupt("return",e.sent);case 3:case"end":return e.stop()}}),e)})))).apply(this,arguments)}var x=n(179),w=u()((function(){return{generatedImg:{borderRadius:"8px"}}}))((function(e){var t=e.classes,n=e.generatedImages,r=e.promptText,o=function(e){var n=e.imgData,r=e.promptText,o=e.index,i="data:image/png;base64,".concat(n),c="".concat(r," ").concat(o),s=r+"_"+o+".jpeg";return a.a.createElement("a",{href:i,alt:c,title:"Download image",download:s},a.a.createElement("img",{src:i,className:t.generatedImg,alt:c,title:"Download image"}))};return a.a.createElement(x.a,{container:!0,alignItems:"center",spacing:3},n.map((function(e,t){return a.a.createElement(x.a,{item:!0,key:t},a.a.createElement(o,{imgData:e,promptText:r,index:++t}))})))})),b=n(184),E=u()((function(){return{inputPrompt:{marginTop:"20px"}}}))((function(e){var t=e.classes,n=e.enterPressedCallback,r=e.disabled,o=e.promptText,i=e.setPromptText;return a.a.createElement(b.a,{className:t.inputPrompt,id:"prompt-input",label:"Text prompt",helperText:"hit Enter to generate images",placeholder:"e.g. an apple on a table",value:o,onChange:function(e){i(e.target.value)},fullWidth:!0,onKeyPress:function(e){"Enter"===e.key&&n(o)},disabled:r})}));n(117);var L=n(44),j=n(76),O=n.n(j),T=(u()((function(){return{inputBackend:{minWidth:"220px"},loadingSpinner:{paddingTop:"20px !important"}}}))((function(e){var t=e.classes,n=e.disabled,o=e.setBackendValidUrl,i=e.isValidBackendEndpoint,s=e.setIsValidBackendEndpoint,u=e.isCheckingBackendEndpoint,l=e.setIsCheckingBackendEndpoint,p=Object(r.useState)(""),f=Object(c.a)(p,2),h=f[0],d=f[1];function m(e){var t;t=e,new RegExp("^(https?:\\/\\/)?((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|((\\d{1,3}\\.){3}\\d{1,3}))(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*(\\?[;&a-z\\d%_.~+=-]*)?(\\#[-a-z\\d_]*)?$","i").test(t)?(l(!0),function(e){return y.apply(this,arguments)}(e).then((function(t){s(t),t&&o(e),l(!1)})).catch((function(){l(!1)}))):s(!1),d(e)}return Object(r.useEffect)((function(){var e=O.a.parse(window.location.search,{ignoreQueryPrefix:!0}).backendUrl;e&&m(e)}),[d]),a.a.createElement(x.a,{container:!0,spacing:1,alignContent:"center"},a.a.createElement(x.a,{item:!0,xs:10},a.a.createElement(b.a,{className:t.inputBackend,fullWidth:!0,id:"standard-basic",label:"Backend URL",value:h,disabled:n,error:!i&&""!==h,helperText:!i&&""!==h&&"No running DALL-E server with this URL",onChange:function(e){return m(e.target.value)}})),u&&a.a.createElement(x.a,{item:!0,className:t.loadingSpinner,xs:2},a.a.createElement(L.PulseLoader,{sizeUnit:"px",size:5,color:"purple",loading:u})))})),["Generating images \ud83d\udc68\ud83c\udffd\u200d\ud83c\udfa8","Doing fancy calculations \u2728"]),k=u()((function(){return{root:{display:"flex",flexDirection:"column",textAlign:"center",marginTop:"80px"},loadingText:{paddingTop:"20px"}}}))((function(e){var t=e.classes,n=e.isLoading,o=Object(r.useState)(0),i=Object(c.a)(o,2),s=i[0],u=i[1];Object(r.useEffect)((function(){var e=setTimeout((function(){s+1<T.length&&u(s+1)}),1e4);return function(){return clearInterval(e)}}),[]);var p=T[s%T.length];return a.a.createElement("div",{className:t.root},a.a.createElement(L.PulseLoader,{sizeUnit:"px",size:20,color:"purple",loading:n}),a.a.createElement(l.a,{className:t.loadingText,variant:"h6"},p))})),S=u()((function(){return{root:{display:"flex",width:"100%",flexDirection:"column",margin:"60px 0px 60px 0px",alignItems:"center",textAlign:"center"},title:{marginBottom:"20px"},playgroundSection:{display:"flex",flex:1,width:"100%",alignItems:"flex-start",justifyContent:"center",marginTop:"20px"},settingsSection:{display:"flex",flexDirection:"column",padding:"1em",maxWidth:"300px"},searchQueryCard:{marginBottom:"20px"},imagesPerQueryControl:{marginTop:"20px"},formControl:{margin:"20px",minWidth:120},gallery:{display:"flex",flex:"1",maxWidth:"50%",height:"100%",justifyContent:"center",alignItems:"flex-start",padding:"1rem"}}}))((function(e){var t=e.classes,n=Object(r.useState)("https://api.imec.ai"),o=Object(c.a)(n,2),i=o[0],s=(o[1],Object(r.useState)(!1)),u=Object(c.a)(s,2),h=u[0],d=u[1],m=Object(r.useState)(!1),g=Object(c.a)(m,2),y=(g[0],g[1],Object(r.useState)(!0)),x=Object(c.a)(y,2),b=x[0],L=(x[1],Object(r.useState)([])),j=Object(c.a)(L,2),O=j[0],T=j[1],S=Object(r.useState)(""),N=Object(c.a)(S,2),P=N[0],_=N[1],I=Object(r.useState)(2),C=Object(c.a)(I,2),D=C[0],B=(C[1],Object(r.useState)(0)),z=Object(c.a)(B,2),A=z[0],G=z[1],W=Object(r.useState)(""),R=Object(c.a)(W,2),U=R[0],F=R[1],q=b&&i;return a.a.createElement("div",{className:t.root},a.a.createElement("div",{className:t.title},a.a.createElement(l.a,{variant:"h3"},"imec.ai demo"," ",a.a.createElement("span",{role:"img","aria-label":"sparks-emoji"},"\u2728"))),a.a.createElement("div",{className:t.playgroundSection},a.a.createElement("div",{className:t.settingsSection},a.a.createElement(p.a,{className:t.searchQueryCard},a.a.createElement(f.a,null,a.a.createElement(E,{enterPressedCallback:function(e){console.log("API call to DALL-E web service with the following prompt ["+e+"]"),_(""),d(!0),function(e,t,n){return v.apply(this,arguments)}(i,e,D).then((function(e){G(e.executionTime),T(e.generatedImgs),d(!1)})).catch((function(e){console.log("Error querying DALL-E service.",e),"Timeout"===e.message?_("Timeout querying DALL-E service (>1min). Consider reducing the images per query or use a stronger backend."):_("Error querying DALL-E service. Check your backend server logs."),d(!1)}))},promptText:U,setPromptText:F,disabled:h||!q}))),0!==A&&a.a.createElement(l.a,{variant:"body2",color:"textSecondary"},"Generation execution time: ",A," sec")),(O.length>0||P||h)&&a.a.createElement("div",{className:t.gallery},P?a.a.createElement(l.a,{variant:"h5",color:"error"},P):h?a.a.createElement(k,{isLoading:h}):a.a.createElement(w,{generatedImages:O,promptText:U}))))}));Boolean("localhost"===window.location.hostname||"[::1]"===window.location.hostname||window.location.hostname.match(/^127(?:\.(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)){3}$/));i.a.render(a.a.createElement(a.a.StrictMode,null,a.a.createElement(S,null)),document.getElementById("root")),"serviceWorker"in navigator&&navigator.serviceWorker.ready.then((function(e){e.unregister()})).catch((function(e){console.error(e.message)}))},82:function(e,t,n){e.exports=n(154)},87:function(e,t,n){}},[[82,1,2]]]);
//# sourceMappingURL=main.2ca53ff8.chunk.js.map