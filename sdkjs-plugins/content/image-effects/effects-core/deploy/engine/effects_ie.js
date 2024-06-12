/* eslint-disable linebreak-style */
/* eslint-disable spaced-comment */
/* eslint-disable max-len */
/* eslint-disable no-unused-vars */
/* eslint-disable new-cap */
/* eslint-disable indent */
/* eslint-disable no-trailing-spaces */

// eslint-disable-next-line prefer-const

(function(window) {
    let isModuleLoaded = false;
    /**
     */
    function onLoadModule() {
        isModuleLoaded = true;
        if (window.ImageEffects) {
            window.ImageEffects.onLoadModule && window.ImageEffects.onLoadModule({
                ApplyEffect: ApplyEffect,
            });
        } else {
            postMessage('module is ready');
        }
    };

    /* Working with Module */

    // correct fetch for desktop application

var printErr = undefined;
var print    = undefined;

var fetch = self.fetch;
var getBinaryPromise = null;

function internal_isLocal()
{
	if (window.navigator && window.navigator.userAgent.toLowerCase().indexOf("ascdesktopeditor") < 0)
		return false;
	if (window.location && window.location.protocol == "file:")
		return true;
	if (window.document && window.document.currentScript && 0 == window.document.currentScript.src.indexOf("file:///"))
		return true;
	return false;
}

if (internal_isLocal())
{
	fetch = undefined; // fetch not support file:/// scheme
	getBinaryPromise = function()
	{
		var wasmPath = "ascdesktop://fonts/" + wasmBinaryFile.substr(8);
		return new Promise(function (resolve, reject)
		{
			var xhr = new XMLHttpRequest();
			xhr.open('GET', wasmPath, true);
			xhr.responseType = 'arraybuffer';

			if (xhr.overrideMimeType)
				xhr.overrideMimeType('text/plain; charset=x-user-defined');
			else
				xhr.setRequestHeader('Accept-Charset', 'x-user-defined');

			xhr.onload = function ()
			{
				if (this.status == 200)
					resolve(new Uint8Array(this.response));
			};
			xhr.send(null);
		});
	}
}
else
{
	getBinaryPromise = function() { return getBinaryPromise2(); }
}


    var ob;function pb(h){var f=0;return function(){return f<h.length?{done:!1,value:h[f++]}:{done:!0}}}function qb(h){var f="undefined"!=typeof Symbol&&Symbol.iterator&&h[Symbol.iterator];return f?f.call(h):{next:pb(h)}}var dd="undefined"!=typeof window&&window===this?this:"undefined"!=typeof global&&null!=global?global:this,Fd="function"==typeof Object.defineProperties?Object.defineProperty:function(h,f,Ka){h!=Array.prototype&&h!=Object.prototype&&(h[f]=Ka.value)};if(!dd)dd=self;
function Gd(h,f){if(f){var Ka=dd;h=h.split(".");for(var Za=0;Za<h.length-1;Za++){var bb=h[Za];bb in Ka||(Ka[bb]={});Ka=Ka[bb]}h=h[h.length-1];Za=Ka[h];f=f(Za);f!=Za&&null!=f&&Fd(Ka,h,{configurable:!0,writable:!0,value:f})}}
Gd("Promise",function(h){function f(f){this.MQf=0;this.Cug=void 0;this.Qie=[];var h=this.slg();try{f(h.resolve,h.reject)}catch(Tb){h.reject(Tb)}}function Ka(){this.FAd=null}function Za(h){return h instanceof f?h:new f(function(f){f(h)})}if(h)return h;Ka.prototype.nJg=function(f){if(null==this.FAd){this.FAd=[];var h=this;this.oJg(function(){h.Ihh()})}this.FAd.push(f)};var bb=dd.setTimeout;Ka.prototype.oJg=function(f){bb(f,0)};Ka.prototype.Ihh=function(){for(;this.FAd&&this.FAd.length;){var f=this.FAd;
this.FAd=[];for(var h=0;h<f.length;++h){var Ka=f[h];f[h]=null;try{Ka()}catch(jb){this.Heh(jb)}}}this.FAd=null};Ka.prototype.Heh=function(f){this.oJg(function(){throw f;})};f.prototype.slg=function(){function f(f){return function(z){Ka||(Ka=!0,f.call(h,z))}}var h=this,Ka=!1;return{resolve:f(this.Cph),reject:f(this.oug)}};f.prototype.Cph=function(h){if(h===this)this.oug(new TypeError("A Promise cannot resolve to itself"));else if(h instanceof f)this.Aqh(h);else{a:switch(typeof h){case "object":var z=
null!=h;break a;case "function":z=!0;break a;default:z=!1}z?this.Bph(h):this.dNg(h)}};f.prototype.Bph=function(f){var h=void 0;try{h=f.then}catch(Tb){this.oug(Tb);return}"function"==typeof h?this.Bqh(h,f):this.dNg(f)};f.prototype.oug=function(f){this.vWg(2,f)};f.prototype.dNg=function(f){this.vWg(1,f)};f.prototype.vWg=function(f,h){if(0!=this.MQf)throw Error("Cannot settle("+f+", "+h+"): Promise already settled in state"+this.MQf);this.MQf=f;this.Cug=h;this.Jhh()};f.prototype.Jhh=function(){if(null!=
this.Qie){for(var f=0;f<this.Qie.length;++f)gb.nJg(this.Qie[f]);this.Qie=null}};var gb=new Ka;f.prototype.Aqh=function(f){var h=this.slg();f.cZf(h.resolve,h.reject)};f.prototype.Bqh=function(f,h){var z=this.slg();try{f.call(h,z.resolve,z.reject)}catch(jb){z.reject(jb)}};f.prototype.then=function(h,Ka){function z(f,h){return"function"==typeof f?function(h){try{gb(f(h))}catch(hc){Ma(hc)}}:h}var gb,Ma,bb=new f(function(f,h){gb=f;Ma=h});this.cZf(z(h,gb),z(Ka,Ma));return bb};f.prototype.catch=function(f){return this.then(void 0,
f)};f.prototype.cZf=function(f,h){function z(){switch(Ka.MQf){case 1:f(Ka.Cug);break;case 2:h(Ka.Cug);break;default:throw Error("Unexpected state: "+Ka.MQf);}}var Ka=this;null==this.Qie?gb.nJg(z):this.Qie.push(z)};f.resolve=Za;f.reject=function(h){return new f(function(f,z){z(h)})};f.race=function(h){return new f(function(f,z){for(var Ka=qb(h),Ma=Ka.next();!Ma.done;Ma=Ka.next())Za(Ma.value).cZf(f,z)})};f.all=function(h){var z=qb(h),Ka=z.next();return Ka.done?Za([]):new f(function(f,h){function Ma(h){return function(z){Ta[h]=
z;gb--;0==gb&&f(Ta)}}var Ta=[],gb=0;do Ta.push(void 0),gb++,Za(Ka.value).cZf(Ma(Ta.length-1),h),Ka=z.next();while(!Ka.done)})};return f});Gd("Array.prototype.fill",function(h){return h?h:function(f,h,Za){var Ka=this.length||0;0>h&&(h=Math.max(0,Ka+h));if(null==Za||Za>Ka)Za=Ka;Za=Number(Za);0>Za&&(Za=Math.max(0,Ka+Za));for(h=Number(h||0);h<Za;h++)this[h]=f;return this}});
function Hd(h,f,Ka){if(null==h)throw new TypeError("The 'this' value for String.prototype."+Ka+" must not be null or undefined");if(f instanceof RegExp)throw new TypeError("First argument to String.prototype."+Ka+" must not be a regular expression");return h+""}Gd("String.prototype.repeat",function(h){return h?h:function(f){var h=Hd(this,null,"repeat");if(0>f||1342177279<f)throw new RangeError("Invalid count value");f|=0;for(var Za="";f;)if(f&1&&(Za+=h),f>>>=1)h+=h;return Za}});
Gd("Number.isFinite",function(h){return h?h:function(f){return"number"!==typeof f?!1:!isNaN(f)&&Infinity!==f&&-Infinity!==f}});Gd("Number.isInteger",function(h){return h?h:function(f){return Number.isFinite(f)?f===Math.floor(f):!1}});Gd("String.prototype.endsWith",function(h){return h?h:function(f,h){var Ka=Hd(this,f,"endsWith");f+="";void 0===h&&(h=Ka.length);h=Math.max(0,Math.min(h|0,Ka.length));for(var bb=f.length;0<bb&&0<h;)if(Ka[--h]!=f[--bb])return!1;return 0>=bb}});
Gd("String.prototype.padStart",function(h){return h?h:function(f,h){var Ka=Hd(this,null,"padStart");f-=Ka.length;h=void 0!==h?String(h):" ";return(0<f&&h?h.repeat(Math.ceil(f/h.length)).substring(0,f):"")+Ka}});function Be(){Be=function(){};dd.Symbol||(dd.Symbol=De)}function Ee(h,f){this.kYg=h;Fd(this,"description",{configurable:!0,writable:!0,value:f})}Ee.prototype.toString=function(){return this.kYg};
var De=function(){function h(Ka){if(this instanceof h)throw new TypeError("Symbol is not a constructor");return new Ee("jscomp_symbol_"+(Ka||"")+"_"+f++,Ka)}var f=0;return h}();function Ng(){Be();var h=dd.Symbol.iterator;h||(h=dd.Symbol.iterator=dd.Symbol("Symbol.iterator"));"function"!=typeof Array.prototype[h]&&Fd(Array.prototype,h,{configurable:!0,writable:!0,value:function(){return Kh(pb(this))}});Ng=function(){}}
function Kh(h){Ng();h={next:h};h[dd.Symbol.iterator]=function(){return this};return h}function qm(h,f){Ng();h instanceof String&&(h+="");var Ka=0,Za={next:function(){if(Ka<h.length){var bb=Ka++;return{value:f(bb,h[bb]),done:!1}}Za.next=function(){return{done:!0,value:void 0}};return Za.next()}};Za[Symbol.iterator]=function(){return Za};return Za}Gd("Array.prototype.values",function(h){return h?h:function(){return qm(this,function(f,h){return h})}});
Gd("Math.sign",function(h){return h?h:function(f){f=Number(f);return 0===f||isNaN(f)?f:0<f?1:-1}});Gd("Array.prototype.keys",function(h){return h?h:function(){return qm(this,function(f){return f})}});function Sm(h,f){return Object.prototype.hasOwnProperty.call(h,f)}
Gd("WeakMap",function(h){function f(f){this.aCf=(z+=Math.random()+1).toString();if(f){f=qb(f);for(var h;!(h=f.next()).done;)h=h.value,this.set(h[0],h[1])}}function Ka(){}function Za(f){Sm(f,gb)||Fd(f,gb,{value:new Ka})}function bb(f){var h=Object[f];h&&(Object[f]=function(f){if(f instanceof Ka)return f;Za(f);return h(f)})}if(function(){if(!h||!Object.seal)return!1;try{var f=Object.seal({}),z=Object.seal({}),Ka=new h([[f,2],[z,3]]);if(2!=Ka.get(f)||3!=Ka.get(z))return!1;Ka.delete(f);Ka.set(z,4);return!Ka.has(f)&&
4==Ka.get(z)}catch(Ma){return!1}}())return h;var gb="$jscomp_hidden_"+Math.random();bb("freeze");bb("preventExtensions");bb("seal");var z=0;f.prototype.set=function(f,h){Za(f);if(!Sm(f,gb))throw Error("WeakMap key fail: "+f);f[gb][this.aCf]=h;return this};f.prototype.get=function(f){return Sm(f,gb)?f[gb][this.aCf]:void 0};f.prototype.has=function(f){return Sm(f,gb)&&Sm(f[gb],this.aCf)};f.prototype.delete=function(f){return Sm(f,gb)&&Sm(f[gb],this.aCf)?delete f[gb][this.aCf]:!1};return f});
Gd("Map",function(h){function f(){var f={};return f.previous=f.next=f.head=f}function Ka(f,h){var z=f.b4c;return Kh(function(){if(z){for(;z.head!=f.b4c;)z=z.previous;for(;z.next!=z.head;)return z=z.next,{done:!1,value:h(z)};z=null}return{done:!0,value:void 0}})}function Za(f,h){var Ka=h&&typeof h;"object"==Ka||"function"==Ka?gb.has(h)?Ka=gb.get(h):(Ka=""+ ++z,gb.set(h,Ka)):Ka="p_"+h;var Ma=f.rsf[Ka];if(Ma&&Sm(f.rsf,Ka))for(f=0;f<Ma.length;f++){var bb=Ma[f];if(h!==h&&bb.key!==bb.key||h===bb.key)return{id:Ka,
list:Ma,index:f,SNb:bb}}return{id:Ka,list:Ma,index:-1,SNb:void 0}}function bb(h){this.rsf={};this.b4c=f();this.size=0;if(h){h=qb(h);for(var z;!(z=h.next()).done;)z=z.value,this.set(z[0],z[1])}}if(function(){if(!h||"function"!=typeof h||!h.prototype.entries||"function"!=typeof Object.seal)return!1;try{var f=Object.seal({x:4}),z=new h(qb([[f,"s"]]));if("s"!=z.get(f)||1!=z.size||z.get({x:4})||z.set({x:4},"t")!=z||2!=z.size)return!1;var Ka=z.entries(),Ma=Ka.next();if(Ma.done||Ma.value[0]!=f||"s"!=Ma.value[1])return!1;
Ma=Ka.next();return Ma.done||4!=Ma.value[0].x||"t"!=Ma.value[1]||!Ka.next().done?!1:!0}catch(Kb){return!1}}())return h;Ng();var gb=new WeakMap;bb.prototype.set=function(f,h){f=0===f?0:f;var z=Za(this,f);z.list||(z.list=this.rsf[z.id]=[]);z.SNb?z.SNb.value=h:(z.SNb={next:this.b4c,previous:this.b4c.previous,head:this.b4c,key:f,value:h},z.list.push(z.SNb),this.b4c.previous.next=z.SNb,this.b4c.previous=z.SNb,this.size++);return this};bb.prototype.delete=function(f){f=Za(this,f);return f.SNb&&f.list?(f.list.splice(f.index,
1),f.list.length||delete this.rsf[f.id],f.SNb.previous.next=f.SNb.next,f.SNb.next.previous=f.SNb.previous,f.SNb.head=null,this.size--,!0):!1};bb.prototype.clear=function(){this.rsf={};this.b4c=this.b4c.previous=f();this.size=0};bb.prototype.has=function(f){return!!Za(this,f).SNb};bb.prototype.get=function(f){return(f=Za(this,f).SNb)&&f.value};bb.prototype.entries=function(){return Ka(this,function(f){return[f.key,f.value]})};bb.prototype.keys=function(){return Ka(this,function(f){return f.key})};
bb.prototype.values=function(){return Ka(this,function(f){return f.value})};bb.prototype.forEach=function(f,h){for(var z=this.entries(),Ma;!(Ma=z.next()).done;)Ma=Ma.value,f.call(h,Ma[1],Ma[0],this)};bb.prototype[Symbol.iterator]=bb.prototype.entries;var z=0;return bb});function Fw(h,f,Ka){h instanceof String&&(h=String(h));for(var Za=h.length,bb=0;bb<Za;bb++){var gb=h[bb];if(f.call(Ka,gb,bb,h))return{dn:bb,Ju:gb}}return{dn:-1,Ju:void 0}}
Gd("Array.prototype.find",function(h){return h?h:function(f,h){return Fw(this,f,h).Ju}});Gd("String.prototype.startsWith",function(h){return h?h:function(f,h){var Ka=Hd(this,f,"startsWith");f+="";var bb=Ka.length,gb=f.length;h=Math.max(0,Math.min(h|0,Ka.length));for(var z=0;z<gb&&h<bb;)if(Ka[h++]!=f[z++])return!1;return z>=gb}});Gd("Object.is",function(h){return h?h:function(f,h){return f===h?0!==f||1/f===1/h:f!==f&&h!==h}});
Gd("Array.prototype.includes",function(h){return h?h:function(f,h){var Ka=this;Ka instanceof String&&(Ka=String(Ka));var bb=Ka.length;h=h||0;for(0>h&&(h=Math.max(h+bb,0));h<bb;h++){var gb=Ka[h];if(gb===f||Object.is(gb,f))return!0}return!1}});Gd("String.prototype.includes",function(h){return h?h:function(f,h){return-1!==Hd(this,f,"includes").indexOf(f,h||0)}});
Gd("Math.tanh",function(h){return h?h:function(f){f=Number(f);if(0===f)return f;var h=Math.exp(-2*Math.abs(f));h=(1-h)/(1+h);return 0>f?-h:h}});Gd("Math.log1p",function(h){return h?h:function(f){f=Number(f);if(.25>f&&-.25<f){for(var h=f,Za=1,bb=f,gb=0,z=1;gb!=bb;)h*=f,z*=-1,bb=(gb=bb)+z*h/++Za;return bb}return Math.log(1+f)}});Gd("Math.expm1",function(h){return h?h:function(f){f=Number(f);if(.25>f&&-.25<f){for(var h=f,Za=1,bb=f,gb=0;gb!=bb;)h*=f/++Za,bb=(gb=bb)+h;return bb}return Math.exp(f)-1}});
Gd("Math.trunc",function(h){return h?h:function(f){f=Number(f);if(isNaN(f)||Infinity===f||-Infinity===f||0===f)return f;var h=Math.floor(Math.abs(f));return 0>f?-h:h}});Gd("Math.log10",function(h){return h?h:function(f){return Math.log(f)/Math.LN10}});Gd("Math.cosh",function(h){if(h)return h;var f=Math.exp;return function(h){h=Number(h);return(f(h)+f(-h))/2}});Gd("Math.sinh",function(h){if(h)return h;var f=Math.exp;return function(h){h=Number(h);return 0===h?h:(f(h)-f(-h))/2}});
Gd("Math.acosh",function(h){return h?h:function(f){f=Number(f);return Math.log(f+Math.sqrt(f*f-1))}});Gd("Math.atanh",function(h){if(h)return h;var f=Math.log1p;return function(h){h=Number(h);return(f(h)-f(-h))/2}});Gd("Math.asinh",function(h){return h?h:function(f){f=Number(f);if(0===f)return f;var h=Math.log(Math.abs(f)+Math.sqrt(f*f+1));return 0>f?-h:h}});Gd("Array.prototype.findIndex",function(h){return h?h:function(f,h){return Fw(this,f,h).dn}});

Math.imul = Math.imul || function(a, b) {
  var ah = (a >>> 16) & 0xffff;
  var al = a & 0xffff;
  var bh = (b >>> 16) & 0xffff;
  var bl = b & 0xffff;
  // сдвиг на 0 бит закрепляет знак в старшей части числа
  // окончательный |0 преобразует беззнаковое значение обратно в знаковое значение
  return ((al * bl) + (((ah * bl + al * bh) << 16) >>> 0)|0);
};

Math.fround = Math.fround || function(x) {
  return new Float32Array([x])[0];
};

Math.clz32 = Math.clz32 || function(value) {
  value = Number(value) >>> 0;
  return value !== 0 ? 31 - Math.floor(Math.log(value + 0.5) / Math.log(2)) : 32;
};

Uint8Array.prototype.copyWithin = Uint8Array.prototype.copyWithin || function(target, start, end) {
    var tmpArray = this.subarray(start, end);
	this.set(tmpArray, target);
	return this;
};


    var Module=typeof Module!="undefined"?Module:{};
var Promise=function(){function noop(){}function bind(fn,thisArg){return function(){fn.apply(thisArg,arguments)}}function Promise(fn){if(!(this instanceof Promise))throw new TypeError("Promises must be constructed via new");if(typeof fn!="function")throw new TypeError("not a function");this._state=0;this._handled=false;this._value=undefined;this._deferreds=[];doResolve(fn,this)}function handle(self,deferred){while(self._state===3)self=self._value;if(self._state===0){self._deferreds.push(deferred);
return}self._handled=true;Promise._immediateFn(function(){var cb=self._state===1?deferred.onFulfilled:deferred.onRejected;if(cb===null){(self._state===1?resolve:reject)(deferred.promise,self._value);return}var ret;try{ret=cb(self._value)}catch(e){reject(deferred.promise,e);return}resolve(deferred.promise,ret)})}function resolve(self,newValue){try{if(newValue===self)throw new TypeError("A promise cannot be resolved with itself.");if(newValue&&(typeof newValue=="object"||typeof newValue=="function")){var then=
newValue.then;if(newValue instanceof Promise){self._state=3;self._value=newValue;finale(self);return}else if(typeof then=="function"){doResolve(bind(then,newValue),self);return}}self._state=1;self._value=newValue;finale(self)}catch(e){reject(self,e)}}function reject(self,newValue){self._state=2;self._value=newValue;finale(self)}function finale(self){if(self._state===2&&self._deferreds.length===0)Promise._immediateFn(function(){if(!self._handled)Promise._unhandledRejectionFn(self._value)});for(var i=
0,len=self._deferreds.length;i<len;i++)handle(self,self._deferreds[i]);self._deferreds=null}function Handler(onFulfilled,onRejected,promise){this.onFulfilled=typeof onFulfilled=="function"?onFulfilled:null;this.onRejected=typeof onRejected=="function"?onRejected:null;this.promise=promise}function doResolve(fn,self){var done=false;try{fn(function(value){if(done)return;done=true;resolve(self,value)},function(reason){if(done)return;done=true;reject(self,reason)})}catch(ex){if(done)return;done=true;reject(self,
ex)}}Promise.prototype["catch"]=function(onRejected){return this.then(null,onRejected)};Promise.prototype.then=function(onFulfilled,onRejected){var prom=new this.constructor(noop);handle(this,new Handler(onFulfilled,onRejected,prom));return prom};Promise.all=function(arr){return new Promise(function(resolve,reject){if(!Array.isArray(arr))return reject(new TypeError("Promise.all accepts an array"));var args=Array.prototype.slice.call(arr);if(args.length===0)return resolve([]);var remaining=args.length;
function res(i,val){try{if(val&&(typeof val=="object"||typeof val=="function")){var then=val.then;if(typeof then=="function"){then.call(val,function(val){res(i,val)},reject);return}}args[i]=val;if(--remaining===0)resolve(args)}catch(ex){reject(ex)}}for(var i=0;i<args.length;i++)res(i,args[i])})};Promise.resolve=function(value){if(value&&typeof value=="object"&&value.constructor==Promise)return value;return new Promise(function(resolve){resolve(value)})};Promise.reject=function(value){return new Promise(function(resolve,
reject){reject(value)})};Promise.race=function(arr){return new Promise(function(resolve,reject){if(!Array.isArray(arr))return reject(new TypeError("Promise.race accepts an array"));for(var i=0,len=arr.length;i<len;i++)Promise.resolve(arr[i]).then(resolve,reject)})};Promise._immediateFn=typeof setImmediate=="function"&&function(fn){setImmediate(fn)}||function(fn){setTimeout(fn,0)};Promise._unhandledRejectionFn=function _unhandledRejectionFn(err){if(typeof console!="undefined"&&console)console.warn("Possible Unhandled Promise Rejection:",
err)};return Promise}();if(typeof Object.assign=="undefined")Object.assign=function(target,source){for(var i=1;i<arguments.length;i++){var source=arguments[i];if(!source)continue;for(var key in source)if(source.hasOwnProperty(key))target[key]=source[key]}return target};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=function(status,toThrow){throw toThrow;};var ENVIRONMENT_IS_WEB=typeof window=="object";
var ENVIRONMENT_IS_WORKER=typeof importScripts=="function";var ENVIRONMENT_IS_NODE=typeof process=="object"&&typeof process.versions=="object"&&typeof process.versions.node=="string";var scriptDirectory="";function locateFile(path){if(Module["locateFile"])return Module["locateFile"](path,scriptDirectory);return scriptDirectory+path}var read_,readAsync,readBinary,setWindowTitle;
if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER)scriptDirectory=self.location.href;else if(typeof document!="undefined"&&document.currentScript)scriptDirectory=document.currentScript.src;if(scriptDirectory.indexOf("blob:")!==0)scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1);else scriptDirectory="";{read_=function(url){try{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText}catch(err$0){var data=
tryParseAsDataURI(url);if(data)return intArrayToString(data);throw err$0;}};if(ENVIRONMENT_IS_WORKER)readBinary=function(url){try{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)}catch(err$1){var data=tryParseAsDataURI(url);if(data)return data;throw err$1;}};readAsync=function(url,onload,onerror){var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=function(){if(xhr.status==
200||xhr.status==0&&xhr.response){onload(xhr.response);return}var data=tryParseAsDataURI(url);if(data){onload(data.buffer);return}onerror()};xhr.onerror=onerror;xhr.send(null)}}setWindowTitle=function(title){return document.title=title}}else;var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.warn.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];
if(Module["quit"])quit_=Module["quit"];var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];var noExitRuntime=Module["noExitRuntime"]||true;
var WebAssembly={Memory:function(opts){this.buffer=new ArrayBuffer(opts["initial"]*65536)},Module:function(binary){},Instance:function(module,info){this.exports=(
// EMSCRIPTEN_START_ASM
function instantiate(V){function c(d){d.set=function(a,b){this[a]=b};d.get=function(a){return this[a]};return d}var e;var f=new Uint8Array(123);for(var a=25;a>=0;--a){f[48+a]=52+a;f[65+a]=a;f[97+a]=26+a}f[43]=62;f[47]=63;function l(m,n,o){var g,h,a=0,i=n,j=o.length,k=n+(j*3>>2)-(o[j-2]=="=")-(o[j-1]=="=");for(;a<j;a+=4){g=f[o.charCodeAt(a+1)];h=f[o.charCodeAt(a+2)];m[i++]=f[o.charCodeAt(a)]<<2|g>>4;if(i<k)m[i++]=g<<4|h>>2;if(i<k)m[i++]=h<<6|f[o.charCodeAt(a+3)]}}function p(q){l(e,1032,"AQ==");l(e,1044,"Ag==");l(e,1056,"Aw==");l(e,1068,"BA==");l(e,1080,"BQ==");l(e,1092,"Bg==");l(e,1104,"Bw==");l(e,1116,"CA==");l(e,1120,"YAZQ")}function T(U){var r=U.a;var s=r.buffer;var t=new Int8Array(s);var u=new Int16Array(s);var v=new Int32Array(s);var w=new Uint8Array(s);var x=new Uint16Array(s);var y=new Uint32Array(s);var z=new Float32Array(s);var A=new Float64Array(s);var B=Math.imul;var C=Math.fround;var D=Math.abs;var E=Math.clz32;var F=Math.min;var G=Math.max;var H=Math.floor;var I=Math.ceil;var J=Math.trunc;var K=Math.sqrt;var L=U.abort;var M=NaN;var N=Infinity;var O=U.b;var P=U.c;var Q=5244512;
// EMSCRIPTEN_START_FUNCS
function ia(a){a=a|0;var b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0;l=Q-16|0;Q=l;a:{b:{c:{d:{e:{f:{g:{h:{i:{j:{k:{if(a>>>0<=244){g=v[282];h=a>>>0<11?16:a+11&-8;c=h>>>3|0;b=g>>>c|0;if(b&3){c=c+((b^-1)&1)|0;a=c<<3;b=a+1168|0;d=v[a+1176>>2];a=v[d+8>>2];l:{if((b|0)==(a|0)){v[282]=wa(c)&g;break l}v[a+12>>2]=b;v[b+8>>2]=a}a=d+8|0;b=c<<3;v[d+4>>2]=b|3;b=b+d|0;v[b+4>>2]=v[b+4>>2]|1;break a}k=v[284];if(k>>>0>=h>>>0){break k}if(b){a=2<<c;a=(0-a|a)&b<<c;b=(0-a&a)-1|0;a=b>>>12&16;c=a;b=b>>>a|0;a=b>>>5&8;c=c|a;b=b>>>a|0;a=b>>>2&4;c=c|a;b=b>>>a|0;a=b>>>1&2;c=c|a;b=b>>>a|0;a=b>>>1&1;d=(c|a)+(b>>>a|0)|0;a=d<<3;b=a+1168|0;e=v[a+1176>>2];a=v[e+8>>2];m:{if((b|0)==(a|0)){g=wa(d)&g;v[282]=g;break m}v[a+12>>2]=b;v[b+8>>2]=a}v[e+4>>2]=h|3;c=e+h|0;a=d<<3;d=a-h|0;v[c+4>>2]=d|1;v[a+e>>2]=d;if(k){b=(k&-8)+1168|0;f=v[287];a=1<<(k>>>3);n:{if(!(a&g)){v[282]=a|g;a=b;break n}a=v[b+8>>2]}v[b+8>>2]=f;v[a+12>>2]=f;v[f+12>>2]=b;v[f+8>>2]=a}a=e+8|0;v[287]=c;v[284]=d;break a}j=v[283];if(!j){break k}b=(0-j&j)-1|0;a=b>>>12&16;c=a;b=b>>>a|0;a=b>>>5&8;c=c|a;b=b>>>a|0;a=b>>>2&4;c=c|a;b=b>>>a|0;a=b>>>1&2;c=c|a;b=b>>>a|0;a=b>>>1&1;c=v[((c|a)+(b>>>a|0)<<2)+1432>>2];f=(v[c+4>>2]&-8)-h|0;b=c;while(1){o:{a=v[b+16>>2];if(!a){a=v[b+20>>2];if(!a){break o}}b=(v[a+4>>2]&-8)-h|0;d=b>>>0<f>>>0;f=d?b:f;c=d?a:c;b=a;continue}break}i=v[c+24>>2];d=v[c+12>>2];if((d|0)!=(c|0)){a=v[c+8>>2];v[a+12>>2]=d;v[d+8>>2]=a;break b}b=c+20|0;a=v[b>>2];if(!a){a=v[c+16>>2];if(!a){break j}b=c+16|0}while(1){e=b;d=a;b=a+20|0;a=v[b>>2];if(a){continue}b=d+16|0;a=v[d+16>>2];if(a){continue}break}v[e>>2]=0;break b}h=-1;if(a>>>0>4294967231){break k}a=a+11|0;h=a&-8;j=v[283];if(!j){break k}f=0-h|0;g=0;p:{if(h>>>0<256){break p}g=31;if(h>>>0>16777215){break p}b=a>>>8|0;a=b+1048320>>>16&8;c=a;b=b<<a;a=b+520192>>>16&4;c=c|a;b=b<<a;a=b+245760>>>16&2;a=((c|a)^14)+(b<<a>>>15|0)|0;g=h>>>a+7&1|a<<1}b=v[(g<<2)+1432>>2];q:{r:{s:{if(!b){a=0;break s}a=0;c=h<<((g|0)==31?0:25-(g>>>1|0)|0);while(1){t:{e=(v[b+4>>2]&-8)-h|0;if(e>>>0>=f>>>0){break t}d=b;f=e;if(e){break t}f=0;a=b;break r}e=v[b+20>>2];b=v[((c>>>29&4)+b|0)+16>>2];a=e?(e|0)==(b|0)?a:e:a;c=c<<1;if(b){continue}break}}if(!(a|d)){d=0;a=2<<g;a=(0-a|a)&j;if(!a){break k}b=(a&0-a)-1|0;a=b>>>12&16;c=a;b=b>>>a|0;a=b>>>5&8;c=c|a;b=b>>>a|0;a=b>>>2&4;c=c|a;b=b>>>a|0;a=b>>>1&2;c=c|a;b=b>>>a|0;a=b>>>1&1;a=v[((c|a)+(b>>>a|0)<<2)+1432>>2]}if(!a){break q}}while(1){b=(v[a+4>>2]&-8)-h|0;c=b>>>0<f>>>0;f=c?b:f;d=c?a:d;b=v[a+16>>2];if(b){a=b}else{a=v[a+20>>2]}if(a){continue}break}}if(!d|v[284]-h>>>0<=f>>>0){break k}g=v[d+24>>2];c=v[d+12>>2];if((d|0)!=(c|0)){a=v[d+8>>2];v[a+12>>2]=c;v[c+8>>2]=a;break c}b=d+20|0;a=v[b>>2];if(!a){a=v[d+16>>2];if(!a){break i}b=d+16|0}while(1){e=b;c=a;b=a+20|0;a=v[b>>2];if(a){continue}b=c+16|0;a=v[c+16>>2];if(a){continue}break}v[e>>2]=0;break c}c=v[284];if(c>>>0>=h>>>0){d=v[287];b=c-h|0;u:{if(b>>>0>=16){v[284]=b;a=d+h|0;v[287]=a;v[a+4>>2]=b|1;v[c+d>>2]=b;v[d+4>>2]=h|3;break u}v[287]=0;v[284]=0;v[d+4>>2]=c|3;a=c+d|0;v[a+4>>2]=v[a+4>>2]|1}a=d+8|0;break a}i=v[285];if(i>>>0>h>>>0){b=i-h|0;v[285]=b;c=v[288];a=c+h|0;v[288]=a;v[a+4>>2]=b|1;v[c+4>>2]=h|3;a=c+8|0;break a}a=0;j=h+47|0;if(v[400]){c=v[402]}else{v[403]=-1;v[404]=-1;v[401]=4096;v[402]=4096;v[400]=l+12&-16^1431655768;v[405]=0;v[393]=0;c=4096}e=j+c|0;f=0-c|0;b=e&f;if(b>>>0<=h>>>0){break a}d=v[392];if(d){c=v[390];g=c+b|0;if(d>>>0<g>>>0|c>>>0>=g>>>0){break a}}if(w[1572]&4){break f}v:{w:{d=v[288];if(d){a=1576;while(1){c=v[a>>2];if(c>>>0<=d>>>0&d>>>0<c+v[a+4>>2]>>>0){break w}a=v[a+8>>2];if(a){continue}break}}c=Z(0);if((c|0)==-1){break g}g=b;d=v[401];a=d-1|0;if(a&c){g=(b-c|0)+(a+c&0-d)|0}if(g>>>0<=h>>>0|g>>>0>2147483646){break g}d=v[392];if(d){a=v[390];f=a+g|0;if(d>>>0<f>>>0|a>>>0>=f>>>0){break g}}a=Z(g);if((c|0)!=(a|0)){break v}break e}g=f&e-i;if(g>>>0>2147483646){break g}c=Z(g);if((c|0)==(v[a>>2]+v[a+4>>2]|0)){break h}a=c}if(!((a|0)==-1|h+48>>>0<=g>>>0)){c=v[402];c=c+(j-g|0)&0-c;if(c>>>0>2147483646){c=a;break e}if((Z(c)|0)!=-1){g=c+g|0;c=a;break e}Z(0-g|0);break g}c=a;if((a|0)!=-1){break e}break g}d=0;break b}c=0;break c}if((c|0)!=-1){break e}}v[393]=v[393]|4}if(b>>>0>2147483646){break d}c=Z(b);a=Z(0);if((c|0)==-1|(a|0)==-1|a>>>0<=c>>>0){break d}g=a-c|0;if(g>>>0<=h+40>>>0){break d}}a=v[390]+g|0;v[390]=a;if(a>>>0>y[391]){v[391]=a}x:{y:{z:{e=v[288];if(e){a=1576;while(1){d=v[a>>2];b=v[a+4>>2];if((d+b|0)==(c|0)){break z}a=v[a+8>>2];if(a){continue}break}break y}a=v[286];if(!(a>>>0<=c>>>0?a:0)){v[286]=c}a=0;v[395]=g;v[394]=c;v[290]=-1;v[291]=v[400];v[397]=0;while(1){d=a<<3;b=d+1168|0;v[d+1176>>2]=b;v[d+1180>>2]=b;a=a+1|0;if((a|0)!=32){continue}break}d=g-40|0;a=c+8&7?-8-c&7:0;b=d-a|0;v[285]=b;a=a+c|0;v[288]=a;v[a+4>>2]=b|1;v[(c+d|0)+4>>2]=40;v[289]=v[404];break x}if(w[a+12|0]&8|d>>>0>e>>>0|c>>>0<=e>>>0){break y}v[a+4>>2]=b+g;a=e+8&7?-8-e&7:0;c=a+e|0;v[288]=c;b=v[285]+g|0;a=b-a|0;v[285]=a;v[c+4>>2]=a|1;v[(b+e|0)+4>>2]=40;v[289]=v[404];break x}if(y[286]>c>>>0){v[286]=c}b=c+g|0;a=1576;A:{B:{C:{D:{E:{F:{while(1){if((b|0)!=v[a>>2]){a=v[a+8>>2];if(a){continue}break F}break}if(!(w[a+12|0]&8)){break E}}a=1576;while(1){b=v[a>>2];if(b>>>0<=e>>>0){f=b+v[a+4>>2]|0;if(f>>>0>e>>>0){break D}}a=v[a+8>>2];continue}}v[a>>2]=c;v[a+4>>2]=v[a+4>>2]+g;j=(c+8&7?-8-c&7:0)+c|0;v[j+4>>2]=h|3;g=b+(b+8&7?-8-b&7:0)|0;i=h+j|0;a=g-i|0;if((e|0)==(g|0)){v[288]=i;a=v[285]+a|0;v[285]=a;v[i+4>>2]=a|1;break B}if(v[287]==(g|0)){v[287]=i;a=v[284]+a|0;v[284]=a;v[i+4>>2]=a|1;v[a+i>>2]=a;break B}f=v[g+4>>2];if((f&3)==1){e=f&-8;G:{if(f>>>0<=255){d=v[g+8>>2];b=f>>>3|0;c=v[g+12>>2];if((c|0)==(d|0)){v[282]=v[282]&wa(b);break G}v[d+12>>2]=c;v[c+8>>2]=d;break G}h=v[g+24>>2];c=v[g+12>>2];H:{if((g|0)!=(c|0)){b=v[g+8>>2];v[b+12>>2]=c;v[c+8>>2]=b;break H}I:{f=g+20|0;b=v[f>>2];if(b){break I}f=g+16|0;b=v[f>>2];if(b){break I}c=0;break H}while(1){d=f;c=b;f=b+20|0;b=v[f>>2];if(b){continue}f=c+16|0;b=v[c+16>>2];if(b){continue}break}v[d>>2]=0}if(!h){break G}d=v[g+28>>2];b=(d<<2)+1432|0;J:{if(v[b>>2]==(g|0)){v[b>>2]=c;if(c){break J}v[283]=v[283]&wa(d);break G}v[h+(v[h+16>>2]==(g|0)?16:20)>>2]=c;if(!c){break G}}v[c+24>>2]=h;b=v[g+16>>2];if(b){v[c+16>>2]=b;v[b+24>>2]=c}b=v[g+20>>2];if(!b){break G}v[c+20>>2]=b;v[b+24>>2]=c}g=e+g|0;f=v[g+4>>2];a=a+e|0}v[g+4>>2]=f&-2;v[i+4>>2]=a|1;v[a+i>>2]=a;if(a>>>0<=255){b=(a&-8)+1168|0;c=v[282];a=1<<(a>>>3);K:{if(!(c&a)){v[282]=a|c;a=b;break K}a=v[b+8>>2]}v[b+8>>2]=i;v[a+12>>2]=i;v[i+12>>2]=b;v[i+8>>2]=a;break B}f=31;if(a>>>0<=16777215){c=a>>>8|0;b=c+1048320>>>16&8;d=b;c=c<<b;b=c+520192>>>16&4;d=d|b;c=c<<b;b=c+245760>>>16&2;b=((d|b)^14)+(c<<b>>>15|0)|0;f=a>>>b+7&1|b<<1}v[i+28>>2]=f;v[i+16>>2]=0;v[i+20>>2]=0;b=(f<<2)+1432|0;d=v[283];c=1<<f;L:{if(!(d&c)){v[283]=c|d;v[b>>2]=i;break L}f=a<<((f|0)==31?0:25-(f>>>1|0)|0);c=v[b>>2];while(1){b=c;if((v[b+4>>2]&-8)==(a|0)){break C}c=f>>>29|0;f=f<<1;d=(c&4)+b|0;c=v[d+16>>2];if(c){continue}break}v[d+16>>2]=i}v[i+24>>2]=b;v[i+12>>2]=i;v[i+8>>2]=i;break B}d=g-40|0;a=c+8&7?-8-c&7:0;b=d-a|0;v[285]=b;a=a+c|0;v[288]=a;v[a+4>>2]=b|1;v[(c+d|0)+4>>2]=40;v[289]=v[404];a=(f+(f-39&7?39-f&7:0)|0)-47|0;d=a>>>0<e+16>>>0?e:a;v[d+4>>2]=27;a=v[397];v[d+16>>2]=v[396];v[d+20>>2]=a;a=v[395];v[d+8>>2]=v[394];v[d+12>>2]=a;v[396]=d+8;v[395]=g;v[394]=c;v[397]=0;a=d+24|0;while(1){v[a+4>>2]=7;b=a+8|0;a=a+4|0;if(b>>>0<f>>>0){continue}break}if((d|0)==(e|0)){break x}v[d+4>>2]=v[d+4>>2]&-2;f=d-e|0;v[e+4>>2]=f|1;v[d>>2]=f;if(f>>>0<=255){b=(f&-8)+1168|0;c=v[282];a=1<<(f>>>3);M:{if(!(c&a)){v[282]=a|c;a=b;break M}a=v[b+8>>2]}v[b+8>>2]=e;v[a+12>>2]=e;v[e+12>>2]=b;v[e+8>>2]=a;break x}a=31;if(f>>>0<=16777215){b=f>>>8|0;a=b+1048320>>>16&8;c=a;b=b<<a;a=b+520192>>>16&4;c=c|a;b=b<<a;a=b+245760>>>16&2;a=((c|a)^14)+(b<<a>>>15|0)|0;a=f>>>a+7&1|a<<1}v[e+28>>2]=a;v[e+16>>2]=0;v[e+20>>2]=0;b=(a<<2)+1432|0;d=v[283];c=1<<a;N:{if(!(d&c)){v[283]=c|d;v[b>>2]=e;break N}a=f<<((a|0)==31?0:25-(a>>>1|0)|0);d=v[b>>2];while(1){b=d;if((f|0)==(v[b+4>>2]&-8)){break A}c=a>>>29|0;a=a<<1;c=(c&4)+b|0;d=v[c+16>>2];if(d){continue}break}v[c+16>>2]=e}v[e+24>>2]=b;v[e+12>>2]=e;v[e+8>>2]=e;break x}a=v[b+8>>2];v[a+12>>2]=i;v[b+8>>2]=i;v[i+24>>2]=0;v[i+12>>2]=b;v[i+8>>2]=a}a=j+8|0;break a}a=v[b+8>>2];v[a+12>>2]=e;v[b+8>>2]=e;v[e+24>>2]=0;v[e+12>>2]=b;v[e+8>>2]=a}a=v[285];if(a>>>0<=h>>>0){break d}b=a-h|0;v[285]=b;c=v[288];a=c+h|0;v[288]=a;v[a+4>>2]=b|1;v[c+4>>2]=h|3;a=c+8|0;break a}v[281]=48;a=0;break a}O:{if(!g){break O}b=v[d+28>>2];a=(b<<2)+1432|0;P:{if(v[a>>2]==(d|0)){v[a>>2]=c;if(c){break P}j=wa(b)&j;v[283]=j;break O}v[g+(v[g+16>>2]==(d|0)?16:20)>>2]=c;if(!c){break O}}v[c+24>>2]=g;a=v[d+16>>2];if(a){v[c+16>>2]=a;v[a+24>>2]=c}a=v[d+20>>2];if(!a){break O}v[c+20>>2]=a;v[a+24>>2]=c}Q:{if(f>>>0<=15){a=f+h|0;v[d+4>>2]=a|3;a=a+d|0;v[a+4>>2]=v[a+4>>2]|1;break Q}v[d+4>>2]=h|3;e=d+h|0;v[e+4>>2]=f|1;v[e+f>>2]=f;if(f>>>0<=255){b=(f&-8)+1168|0;c=v[282];a=1<<(f>>>3);R:{if(!(c&a)){v[282]=a|c;a=b;break R}a=v[b+8>>2]}v[b+8>>2]=e;v[a+12>>2]=e;v[e+12>>2]=b;v[e+8>>2]=a;break Q}a=31;if(f>>>0<=16777215){b=f>>>8|0;a=b+1048320>>>16&8;c=a;b=b<<a;a=b+520192>>>16&4;c=c|a;b=b<<a;a=b+245760>>>16&2;a=((c|a)^14)+(b<<a>>>15|0)|0;a=f>>>a+7&1|a<<1}v[e+28>>2]=a;v[e+16>>2]=0;v[e+20>>2]=0;b=(a<<2)+1432|0;S:{c=1<<a;T:{if(!(c&j)){v[283]=c|j;v[b>>2]=e;break T}a=f<<((a|0)==31?0:25-(a>>>1|0)|0);h=v[b>>2];while(1){b=h;if((v[b+4>>2]&-8)==(f|0)){break S}c=a>>>29|0;a=a<<1;c=(c&4)+b|0;h=v[c+16>>2];if(h){continue}break}v[c+16>>2]=e}v[e+24>>2]=b;v[e+12>>2]=e;v[e+8>>2]=e;break Q}a=v[b+8>>2];v[a+12>>2]=e;v[b+8>>2]=e;v[e+24>>2]=0;v[e+12>>2]=b;v[e+8>>2]=a}a=d+8|0;break a}U:{if(!i){break U}b=v[c+28>>2];a=(b<<2)+1432|0;V:{if(v[a>>2]==(c|0)){v[a>>2]=d;if(d){break V}v[283]=wa(b)&j;break U}v[i+(v[i+16>>2]==(c|0)?16:20)>>2]=d;if(!d){break U}}v[d+24>>2]=i;a=v[c+16>>2];if(a){v[d+16>>2]=a;v[a+24>>2]=d}a=v[c+20>>2];if(!a){break U}v[d+20>>2]=a;v[a+24>>2]=d}W:{if(f>>>0<=15){a=f+h|0;v[c+4>>2]=a|3;a=a+c|0;v[a+4>>2]=v[a+4>>2]|1;break W}v[c+4>>2]=h|3;d=c+h|0;v[d+4>>2]=f|1;v[d+f>>2]=f;if(k){b=(k&-8)+1168|0;e=v[287];a=1<<(k>>>3);X:{if(!(a&g)){v[282]=a|g;a=b;break X}a=v[b+8>>2]}v[b+8>>2]=e;v[a+12>>2]=e;v[e+12>>2]=b;v[e+8>>2]=a}v[287]=d;v[284]=f}a=c+8|0}Q=l+16|0;return a|0}function W(a){a=a|0;var b=0,c=0,d=0,e=0,f=0,g=0,h=0;a:{if(!a){break a}d=a-8|0;b=v[a-4>>2];a=b&-8;f=d+a|0;b:{if(b&1){break b}if(!(b&3)){break a}b=v[d>>2];d=d-b|0;if(d>>>0<y[286]){break a}a=a+b|0;if(v[287]!=(d|0)){if(b>>>0<=255){e=v[d+8>>2];b=b>>>3|0;c=v[d+12>>2];if((c|0)==(e|0)){v[282]=v[282]&wa(b);break b}v[e+12>>2]=c;v[c+8>>2]=e;break b}h=v[d+24>>2];b=v[d+12>>2];c:{if((d|0)!=(b|0)){c=v[d+8>>2];v[c+12>>2]=b;v[b+8>>2]=c;break c}d:{e=d+20|0;c=v[e>>2];if(c){break d}e=d+16|0;c=v[e>>2];if(c){break d}b=0;break c}while(1){g=e;b=c;e=b+20|0;c=v[e>>2];if(c){continue}e=b+16|0;c=v[b+16>>2];if(c){continue}break}v[g>>2]=0}if(!h){break b}e=v[d+28>>2];c=(e<<2)+1432|0;e:{if(v[c>>2]==(d|0)){v[c>>2]=b;if(b){break e}v[283]=v[283]&wa(e);break b}v[h+(v[h+16>>2]==(d|0)?16:20)>>2]=b;if(!b){break b}}v[b+24>>2]=h;c=v[d+16>>2];if(c){v[b+16>>2]=c;v[c+24>>2]=b}c=v[d+20>>2];if(!c){break b}v[b+20>>2]=c;v[c+24>>2]=b;break b}b=v[f+4>>2];if((b&3)!=3){break b}v[284]=a;v[f+4>>2]=b&-2;v[d+4>>2]=a|1;v[a+d>>2]=a;return}if(d>>>0>=f>>>0){break a}b=v[f+4>>2];if(!(b&1)){break a}f:{if(!(b&2)){if(v[288]==(f|0)){v[288]=d;a=v[285]+a|0;v[285]=a;v[d+4>>2]=a|1;if(v[287]!=(d|0)){break a}v[284]=0;v[287]=0;return}if(v[287]==(f|0)){v[287]=d;a=v[284]+a|0;v[284]=a;v[d+4>>2]=a|1;v[a+d>>2]=a;return}a=(b&-8)+a|0;g:{if(b>>>0<=255){e=v[f+8>>2];b=b>>>3|0;c=v[f+12>>2];if((c|0)==(e|0)){v[282]=v[282]&wa(b);break g}v[e+12>>2]=c;v[c+8>>2]=e;break g}h=v[f+24>>2];b=v[f+12>>2];h:{if((f|0)!=(b|0)){c=v[f+8>>2];v[c+12>>2]=b;v[b+8>>2]=c;break h}i:{e=f+20|0;c=v[e>>2];if(c){break i}e=f+16|0;c=v[e>>2];if(c){break i}b=0;break h}while(1){g=e;b=c;e=b+20|0;c=v[e>>2];if(c){continue}e=b+16|0;c=v[b+16>>2];if(c){continue}break}v[g>>2]=0}if(!h){break g}e=v[f+28>>2];c=(e<<2)+1432|0;j:{if(v[c>>2]==(f|0)){v[c>>2]=b;if(b){break j}v[283]=v[283]&wa(e);break g}v[h+(v[h+16>>2]==(f|0)?16:20)>>2]=b;if(!b){break g}}v[b+24>>2]=h;c=v[f+16>>2];if(c){v[b+16>>2]=c;v[c+24>>2]=b}c=v[f+20>>2];if(!c){break g}v[b+20>>2]=c;v[c+24>>2]=b}v[d+4>>2]=a|1;v[a+d>>2]=a;if(v[287]!=(d|0)){break f}v[284]=a;return}v[f+4>>2]=b&-2;v[d+4>>2]=a|1;v[a+d>>2]=a}if(a>>>0<=255){b=(a&-8)+1168|0;c=v[282];a=1<<(a>>>3);k:{if(!(c&a)){v[282]=a|c;a=b;break k}a=v[b+8>>2]}v[b+8>>2]=d;v[a+12>>2]=d;v[d+12>>2]=b;v[d+8>>2]=a;return}e=31;if(a>>>0<=16777215){c=a>>>8|0;b=c+1048320>>>16&8;e=b;c=c<<b;b=c+520192>>>16&4;e=e|b;c=c<<b;b=c+245760>>>16&2;b=((e|b)^14)+(c<<b>>>15|0)|0;e=a>>>b+7&1|b<<1}v[d+28>>2]=e;v[d+16>>2]=0;v[d+20>>2]=0;g=(e<<2)+1432|0;l:{m:{c=v[283];b=1<<e;n:{if(!(c&b)){v[283]=b|c;v[g>>2]=d;v[d+24>>2]=g;break n}e=a<<((e|0)==31?0:25-(e>>>1|0)|0);b=v[g>>2];while(1){c=b;if((v[b+4>>2]&-8)==(a|0)){break m}b=e>>>29|0;e=e<<1;g=c+(b&4)|0;b=v[g+16>>2];if(b){continue}break}v[g+16>>2]=d;v[d+24>>2]=c}v[d+12>>2]=d;v[d+8>>2]=d;break l}a=v[c+8>>2];v[a+12>>2]=d;v[c+8>>2]=d;v[d+24>>2]=0;v[d+12>>2]=c;v[d+8>>2]=a}a=v[290]-1|0;v[290]=a?a:-1}}function fa(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0;c=Q+-64|0;Q=c;i=v[a+4>>2];a:{if(!i){break a}j=v[b+4>>2];if((j|0)<=0){break a}k=v[b+8>>2];if((k|0)<=0){break a}h=+v[a+4>>2]/100;while(1){a=0;while(1){v[c+8>>2]=Y(b,g,a);aa(c+16|0,c+8|0);f=v[c+36>>2];v[c+56>>2]=v[c+32>>2];v[c+60>>2]=f;f=v[c+28>>2];v[c+48>>2]=v[c+24>>2];v[c+52>>2]=f;f=v[c+20>>2];v[c+40>>2]=v[c+16>>2];v[c+44>>2]=f;d=h+A[c+40>>3];A[c+40>>3]=d;b:{if(!(d>1)){if(!(d<0)){break b}v[c+40>>2]=0;v[c+44>>2]=0;break b}v[c+40>>2]=0;v[c+44>>2]=1072693248}e=h+A[c+48>>3];A[c+48>>3]=e;d=1;c:{if(!(e>1)){d=0;if(!(e<0)){break c}}A[c+48>>3]=d}e=h+A[c+56>>3];A[c+56>>3]=e;d=1;d:{if(!(e>1)){d=0;if(!(e<0)){break d}}A[c+56>>3]=d}v[c+16>>2]=Y(b,g,a);$(c+16|0,c+40|0);a=a+1|0;if((k|0)!=(a|0)){continue}break}g=g+1|0;if((j|0)!=(g|0)){continue}break}}Q=c- -64|0;return(i|0)!=0|0}function da(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0;j=v[b+4>>2];k=v[b+8>>2];if(!((j|0)<=0|(k|0)<=0)){while(1){f=0;while(1){a=v[b>>2]+(B(v[b+8>>2],g)+f<<2)|0;d=+w[a+2|0];h=+w[a|0];i=+w[a+1|0];e=F(d*.131+(h*.272+i*.534),255);a:{if(e<4294967296&e>=0){c=~~e>>>0;break a}c=0}t[a+2|0]=c;e=F(d*.168+(h*.349+i*.686),255);b:{if(e<4294967296&e>=0){c=~~e>>>0;break b}c=0}t[a+1|0]=c;d=F(d*.189+(h*.393+i*.769),255);c:{if(d<4294967296&d>=0){c=~~d>>>0;break c}c=0}t[a|0]=c;f=f+1|0;if((k|0)!=(f|0)){continue}break}g=g+1|0;if((j|0)!=(g|0)){continue}break}}return 1}function ca(a,b){var c=0,d=0,e=0,f=0,g=0;c=A[b+16>>3];e=A[b+8>>3];g=A[b>>3];d=e>g?e:g;d=c>d?c:d;f=e<g?e:g;f=c<f?c:f;b=0;a:{b:{if(d==f){break b}c:{d:{if(d!=g){break d}if(c<=e){c=(e-c)/(d-f)*60;if(!(D(c)<2147483648)){break c}b=~~c;break b}if(!(c>e)){break d}c=(e-c)/(d-f)*60+360;if(!(D(c)<2147483648)){break c}b=~~c;break b}if(d==e){c=(c-g)/(d-f)*60+120;if(!(D(c)<2147483648)){break c}b=~~c;break b}if(c!=d){break a}c=(g-e)/(d-f)*60+240;if(!(D(c)<2147483648)){break c}b=~~c;break b}b=-2147483648}u[a>>1]=b}A[a+16>>3]=d;A[a+8>>3]=d!=0?1-f/d:0}function ha(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0;c=Q+-64|0;Q=c;h=v[a+4>>2];a:{if(!h){break a}i=v[b+4>>2];if((i|0)<=0){break a}j=v[b+8>>2];if((j|0)<=0){break a}while(1){e=0;while(1){v[c+8>>2]=Y(b,f,e);g=c+16|0;k=c+8|0;aa(g,k);l=c+40|0;ca(l,g);d=x[c+40>>1]+x[a+4>>1]|0;d=d<<16>>16>360?d-360|0:d;u[c+40>>1]=d<<16>>16<0?d+360|0:d;v[c+8>>2]=Y(b,f,e);ba(g,l);$(k,g);e=e+1|0;if((j|0)!=(e|0)){continue}break}f=f+1|0;if((i|0)!=(f|0)){continue}break}}Q=c- -64|0;return(h|0)!=0|0}function ba(a,b){var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0;i=x[b>>1];j=(i<<16>>16)/60|0;f=A[b+16>>3];e=f*100;c=(100-A[b+8>>3]*100)*f;d=(e-c)*+(i-B(j,60)<<16>>16)/60;g=e-d;h=c+d;a:{b:{switch((j<<16>>16)%6|0){default:d=A[a+8>>3];e=A[a>>3];c=A[a+16>>3];break a;case 0:d=h;break a;case 1:d=f*100;e=g;break a;case 2:d=f*100;e=c;c=h;break a;case 3:d=g;e=c;c=f*100;break a;case 4:d=c;e=h;c=f*100;break a;case 5:break b}}A[a>>3]=e;d=c;c=g}A[a+16>>3]=c/100;A[a+8>>3]=d/100;A[a>>3]=e/100}function ka(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0;c=Q+-64|0;Q=c;g=v[a+4>>2];a:{if(!g){break a}h=v[b+4>>2];if((h|0)<=0){break a}i=v[b+8>>2];if((i|0)<=0){break a}while(1){d=0;while(1){v[c+8>>2]=Y(b,e,d);f=c+16|0;j=c+8|0;aa(f,j);k=c+40|0;ca(k,f);A[c+56>>3]=G(F(A[c+56>>3]+ +v[a+4>>2]/100,1),0);v[c+8>>2]=Y(b,e,d);ba(f,k);$(j,f);d=d+1|0;if((i|0)!=(d|0)){continue}break}e=e+1|0;if((h|0)!=(e|0)){continue}break}}Q=c- -64|0;return(g|0)!=0|0}function ja(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0;c=Q+-64|0;Q=c;g=v[a+4>>2];a:{if(!g){break a}h=v[b+4>>2];if((h|0)<=0){break a}i=v[b+8>>2];if((i|0)<=0){break a}while(1){d=0;while(1){v[c+8>>2]=Y(b,e,d);f=c+16|0;j=c+8|0;aa(f,j);k=c+40|0;ca(k,f);A[c+48>>3]=G(F(A[c+48>>3]+ +v[a+4>>2]/100,1),0);v[c+8>>2]=Y(b,e,d);ba(f,k);$(j,f);d=d+1|0;if((i|0)!=(d|0)){continue}break}e=e+1|0;if((h|0)!=(e|0)){continue}break}}Q=c- -64|0;return(g|0)!=0|0}function ea(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0;h=v[b+4>>2];a=v[b+8>>2];if(!((h|0)<=0|(a|0)<=0)){while(1){e=0;while(1){c=255;d=v[b>>2]+(B(v[b+8>>2],g)+e<<2)|0;i=+w[d+2|0]*.11+(+w[d|0]*.3+ +w[d+1|0]*.59);a:{if(i>255){break a}c=i;if(!(c<0)){break a}c=0}if(c<4294967296&c>=0){f=~~c>>>0}else{f=0}t[d+2|0]=f;t[d+1|0]=f;t[d|0]=f;e=e+1|0;if((e|0)!=(a|0)){continue}break}g=g+1|0;if((h|0)!=(g|0)){continue}break}}return 1}function _(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0;c=v[a+4>>2];a:{if(!c|y[a+8>>2]>2){break a}d=1;f=v[b+4>>2];if((f|0)<=0){break a}g=v[b+8>>2];if((g|0)<=0){break a}i=(B(c,255)|0)/100|0;c=0;while(1){d=0;while(1){h=v[a+8>>2]+(v[b>>2]+(B(v[b+8>>2],c)+d<<2)|0)|0;e=w[h|0]+i|0;e=(e|0)>0?e:0;t[h|0]=e>>>0<255?e:255;d=d+1|0;if((g|0)!=(d|0)){continue}break}d=1;c=c+1|0;if((f|0)!=(c|0)){continue}break}}return d|0}function ga(a,b){a=a|0;b=b|0;var c=0,d=0,e=0,f=0,g=0,h=0,i=0;c=Q-16|0;Q=c;e=v[a+4>>2];a:{if(!e){break a}f=v[b+4>>2];if((f|0)<=0){break a}g=v[b+8>>2];if((g|0)<=0){break a}h=(B(v[a+4>>2],255)|0)/-100|0;while(1){a=0;while(1){v[c+8>>2]=Y(b,d,a);i=w[v[c+8>>2]+3|0];v[c>>2]=Y(b,d,a);t[v[c>>2]+3|0]=h+i;a=a+1|0;if((g|0)!=(a|0)){continue}break}d=d+1|0;if((f|0)!=(d|0)){continue}break}}Q=c+16|0;return(e|0)!=0|0}function $(a,b){var c=0,d=0,e=0;e=v[a>>2];c=A[b>>3]*255;a:{if(c<4294967296&c>=0){d=~~c>>>0;break a}d=0}t[e|0]=d;e=v[a>>2];c=A[b+8>>3]*255;b:{if(c<4294967296&c>=0){d=~~c>>>0;break b}d=0}t[e+1|0]=d;d=v[a>>2];c=A[b+16>>3]*255;c:{if(c<4294967296&c>=0){a=~~c>>>0;break c}a=0}t[d+2|0]=a}function aa(a,b){v[a>>2]=0;v[a+4>>2]=0;v[a+16>>2]=0;v[a+20>>2]=0;v[a+8>>2]=0;v[a+12>>2]=0;b=v[b>>2];A[a>>3]=+w[b|0]/255;A[a+8>>3]=+w[b+1|0]/255;A[a+16>>3]=+w[b+2|0]/255}function ra(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;e=Q-16|0;Q=e;f=X();v[f+8>>2]=c;v[f+4>>2]=b;v[f>>2]=a;v[e+8>>2]=0;v[e+4>>2]=d;v[e>>2]=1092;_(e,f);W(f);Q=e+16|0}function qa(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;e=Q-16|0;Q=e;f=X();v[f+8>>2]=c;v[f+4>>2]=b;v[f>>2]=a;v[e+8>>2]=1;v[e+4>>2]=d;v[e>>2]=1092;_(e,f);W(f);Q=e+16|0}function pa(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;e=Q-16|0;Q=e;f=X();v[f+8>>2]=c;v[f+4>>2]=b;v[f>>2]=a;v[e+8>>2]=2;v[e+4>>2]=d;v[e>>2]=1092;_(e,f);W(f);Q=e+16|0}function Z(a){var b=0,c=0;b=v[280];c=a+3&-4;a=b+c|0;a:{if(a>>>0<=b>>>0?c:0){break a}if(a>>>0>S()<<16>>>0){if(!(O(a|0)|0)){break a}}v[280]=a;return b}v[281]=48;return-1}function va(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;f=Q-16|0;Q=f;e=X();v[e+8>>2]=c;v[e+4>>2]=b;v[e>>2]=a;a=f+8|0;v[a+4>>2]=d;v[a>>2]=1044;ja(a,e);W(e);Q=f+16|0}function ua(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;f=Q-16|0;Q=f;e=X();v[e+8>>2]=c;v[e+4>>2]=b;v[e>>2]=a;a=f+8|0;v[a+4>>2]=d;v[a>>2]=1068;ga(a,e);W(e);Q=f+16|0}function ta(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;f=Q-16|0;Q=f;e=X();v[e+8>>2]=c;v[e+4>>2]=b;v[e>>2]=a;a=f+8|0;v[a+4>>2]=d;v[a>>2]=1056;ha(a,e);W(e);Q=f+16|0}function sa(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;f=Q-16|0;Q=f;e=X();v[e+8>>2]=c;v[e+4>>2]=b;v[e>>2]=a;a=f+8|0;v[a+4>>2]=d;v[a>>2]=1080;fa(a,e);W(e);Q=f+16|0}function ma(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0,f=0;f=Q-16|0;Q=f;e=X();v[e+8>>2]=c;v[e+4>>2]=b;v[e>>2]=a;a=f+8|0;v[a+4>>2]=d;v[a>>2]=1032;ka(a,e);W(e);Q=f+16|0}function oa(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0;e=Q-16|0;Q=e;if(d){d=X();v[d+8>>2]=c;v[d+4>>2]=b;v[d>>2]=a;a=e+8|0;v[a>>2]=1104;ea(a,d);W(d)}Q=e+16|0}function na(a,b,c,d){a=a|0;b=b|0;c=c|0;d=d|0;var e=0;e=Q-16|0;Q=e;if(d){d=X();v[d+8>>2]=c;v[d+4>>2]=b;v[d>>2]=a;a=e+8|0;v[a>>2]=1116;da(a,d);W(d)}Q=e+16|0}function X(){var a=0;a:{while(1){a=ia(12);if(a){break a}a=v[406];if(a){R[a|0]();continue}break}P();L()}return a}function Y(a,b,c){var d=0;d=Q-16|0;Q=d;v[d+8>>2]=v[a>>2]+(B(v[a+8>>2],b)+c<<2);Q=d+16|0;return v[d+8>>2]}function wa(a){var b=0;b=a&31;a=0-a&31;return(-1>>>b&-2)<<b|(-1<<a&-2)>>>a}
function la(){}
// EMSCRIPTEN_END_FUNCS
e=w;p(U);var R=c([null,ka,ja,ha,ga,fa,_,ea,da]);function S(){return s.byteLength/65536|0}return{"d":la,"e":R,"f":ma,"g":va,"h":ua,"i":ta,"j":sa,"k":ra,"l":qa,"m":pa,"n":oa,"o":na,"p":ia,"q":W}}return T(V)}
// EMSCRIPTEN_END_ASM




)(asmLibraryArg)},instantiate:function(binary,info){return{then:function(ok){var module=new WebAssembly.Module(binary);ok({"instance":new WebAssembly.Instance(module)})}}},RuntimeError:Error};wasmBinary=[];if(typeof WebAssembly!="object")abort("no native wasm support detected");var wasmMemory;var ABORT=false;
var EXITSTATUS;function assert(condition,text){if(!condition)abort(text)}var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;
function updateGlobalBufferAndViews(buf){buffer=buf;Module["HEAP8"]=HEAP8=new Int8Array(buf);Module["HEAP16"]=HEAP16=new Int16Array(buf);Module["HEAP32"]=HEAP32=new Int32Array(buf);Module["HEAPU8"]=HEAPU8=new Uint8Array(buf);Module["HEAPU16"]=HEAPU16=new Uint16Array(buf);Module["HEAPU32"]=HEAPU32=new Uint32Array(buf);Module["HEAPF32"]=HEAPF32=new Float32Array(buf);Module["HEAPF64"]=HEAPF64=new Float64Array(buf)}var INITIAL_MEMORY=Module["INITIAL_MEMORY"]||134217728;
if(Module["wasmMemory"])wasmMemory=Module["wasmMemory"];else wasmMemory=new WebAssembly.Memory({"initial":INITIAL_MEMORY/65536,"maximum":INITIAL_MEMORY/65536});if(wasmMemory)buffer=wasmMemory.buffer;INITIAL_MEMORY=buffer.byteLength;updateGlobalBufferAndViews(buffer);var wasmTable;var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[onLoadModule];var runtimeInitialized=false;
function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length)addOnPreRun(Module["preRun"].shift())}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}
function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length)addOnPostRun(Module["postRun"].shift())}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}
if(!Math.imul||Math.imul(4294967295,5)!==-5)Math.imul=function imul(a,b){var ah=a>>>16;var al=a&65535;var bh=b>>>16;var bl=b&65535;return al*bl+(ah*bl+al*bh<<16)|0};if(!Math.fround){var froundBuffer=new Float32Array(1);Math.fround=function(x){froundBuffer[0]=x;return froundBuffer[0]}}if(!Math.clz32)Math.clz32=function(x){var n=32;var y=x>>16;if(y){n-=16;x=y}y=x>>8;if(y){n-=8;x=y}y=x>>4;if(y){n-=4;x=y}y=x>>2;if(y){n-=2;x=y}y=x>>1;if(y)return n-2;return n-x};
if(!Math.trunc)Math.trunc=function(x){return x<0?Math.ceil(x):Math.floor(x)};var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"])Module["monitorRunDependencies"](runDependencies)}
function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"])Module["monitorRunDependencies"](runDependencies);if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}
function abort(what){{if(Module["onAbort"])Module["onAbort"](what)}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -sASSERTIONS for more info.";var e=new WebAssembly.RuntimeError(what);throw e;}var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return filename.startsWith(dataURIPrefix)}var wasmBinaryFile;wasmBinaryFile="effects.wasm";if(!isDataURI(wasmBinaryFile))wasmBinaryFile=locateFile(wasmBinaryFile);
function getBinary(file){try{if(file==wasmBinaryFile&&wasmBinary)return new Uint8Array(wasmBinary);var binary=tryParseAsDataURI(file);if(binary)return binary;if(readBinary)return readBinary(file);else throw"both async and sync fetching of the wasm failed";}catch(err$2){abort(err$2)}}
function getBinaryPromise(){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER))if(typeof fetch=="function")return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"])throw"failed to load wasm binary file at '"+wasmBinaryFile+"'";return response["arrayBuffer"]()}).catch(function(){return getBinary(wasmBinaryFile)});return Promise.resolve().then(function(){return getBinary(wasmBinaryFile)})}
function createWasm(){var info={"a":asmLibraryArg};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;wasmTable=Module["asm"]["e"];addOnInit(Module["asm"]["d"]);removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}function instantiateArrayBuffer(receiver){return getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(function(instance){return instance}).then(receiver,
function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}function instantiateAsync(){if(!wasmBinary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch=="function")return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){var result=WebAssembly.instantiateStreaming(response,info);return result.then(receiveInstantiationResult,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");
return instantiateArrayBuffer(receiveInstantiationResult)})});else return instantiateArrayBuffer(receiveInstantiationResult)}if(Module["instantiateWasm"])try{var exports=Module["instantiateWasm"](info,receiveInstance);return exports}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}instantiateAsync();return{}}function callRuntimeCallbacks(callbacks){while(callbacks.length>0)callbacks.shift()(Module)}
function intArrayToString(array){var ret=[];for(var i=0;i<array.length;i++){var chr=array[i];if(chr>255){if(ASSERTIONS)assert(false,"Character code "+chr+" ("+String.fromCharCode(chr)+")  at offset "+i+" not in 0x00-0xFF.");chr&=255}ret.push(String.fromCharCode(chr))}return ret.join("")}function _abort(){abort("")}function abortOnCannotGrowMemory(requestedSize){abort("OOM")}
function _emscripten_resize_heap(requestedSize){var oldSize=HEAPU8.length;requestedSize=requestedSize>>>0;abortOnCannotGrowMemory(requestedSize)}var ASSERTIONS=false;
var decodeBase64=typeof atob=="function"?atob:function(input){var keyStr="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";var output="";var chr1,chr2,chr3;var enc1,enc2,enc3,enc4;var i=0;input=input.replace(/[^A-Za-z0-9\+\/=]/g,"");do{enc1=keyStr.indexOf(input.charAt(i++));enc2=keyStr.indexOf(input.charAt(i++));enc3=keyStr.indexOf(input.charAt(i++));enc4=keyStr.indexOf(input.charAt(i++));chr1=enc1<<2|enc2>>4;chr2=(enc2&15)<<4|enc3>>2;chr3=(enc3&3)<<6|enc4;output=output+String.fromCharCode(chr1);
if(enc3!==64)output=output+String.fromCharCode(chr2);if(enc4!==64)output=output+String.fromCharCode(chr3)}while(i<input.length);return output};function intArrayFromBase64(s){try{var decoded=decodeBase64(s);var bytes=new Uint8Array(decoded.length);for(var i=0;i<decoded.length;++i)bytes[i]=decoded.charCodeAt(i);return bytes}catch(_){throw new Error("Converting base64 string to bytes failed.");}}
function tryParseAsDataURI(filename){if(!isDataURI(filename))return;return intArrayFromBase64(filename.slice(dataURIPrefix.length))}var asmLibraryArg={"c":_abort,"b":_emscripten_resize_heap,"a":wasmMemory};var asm=createWasm();var ___wasm_call_ctors=Module["___wasm_call_ctors"]=function(){return(___wasm_call_ctors=Module["___wasm_call_ctors"]=Module["asm"]["d"]).apply(null,arguments)};
var _change_brightness=Module["_change_brightness"]=function(){return(_change_brightness=Module["_change_brightness"]=Module["asm"]["f"]).apply(null,arguments)};var _change_saturation=Module["_change_saturation"]=function(){return(_change_saturation=Module["_change_saturation"]=Module["asm"]["g"]).apply(null,arguments)};var _change_transparency=Module["_change_transparency"]=function(){return(_change_transparency=Module["_change_transparency"]=Module["asm"]["h"]).apply(null,arguments)};
var _change_hue=Module["_change_hue"]=function(){return(_change_hue=Module["_change_hue"]=Module["asm"]["i"]).apply(null,arguments)};var _change_lighting=Module["_change_lighting"]=function(){return(_change_lighting=Module["_change_lighting"]=Module["asm"]["j"]).apply(null,arguments)};var _change_intension_red=Module["_change_intension_red"]=function(){return(_change_intension_red=Module["_change_intension_red"]=Module["asm"]["k"]).apply(null,arguments)};
var _change_intension_green=Module["_change_intension_green"]=function(){return(_change_intension_green=Module["_change_intension_green"]=Module["asm"]["l"]).apply(null,arguments)};var _change_intension_blue=Module["_change_intension_blue"]=function(){return(_change_intension_blue=Module["_change_intension_blue"]=Module["asm"]["m"]).apply(null,arguments)};var _shades_of_gray=Module["_shades_of_gray"]=function(){return(_shades_of_gray=Module["_shades_of_gray"]=Module["asm"]["n"]).apply(null,arguments)};
var _sepia=Module["_sepia"]=function(){return(_sepia=Module["_sepia"]=Module["asm"]["o"]).apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return(_malloc=Module["_malloc"]=Module["asm"]["p"]).apply(null,arguments)};var _free=Module["_free"]=function(){return(_free=Module["_free"]=Module["asm"]["q"]).apply(null,arguments)};var calledRun;dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};
function run(args){args=args||arguments_;if(runDependencies>0)return;preRun();if(runDependencies>0)return;function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else doRun()}
if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0)Module["preInit"].pop()()}run();

    /**
     * @param {Module} module
     * @param {Number} length
     * @return {Number}
     * Allocate memory for wasm, returns pointer
     */
     function allocateMemory(module, length) {
        const ptr = module._malloc(length);
        return ptr;
    }

    /**
     * @param {Module} module
     * @param {Number} ptr
     * @return {void}
     * Free memory
     */
    function freeMemory(module, ptr) {
        module._free(ptr);
    }

    /**
     * @param {Module} module
     * @param {Uint8Array} data
     * @param {Number} ptr
     * @return {void}
     * Fill memory
     */
    function setMemory(module, data, ptr) {
        module.HEAP8.set(data, ptr);
    }

    const effect = {
        /* One Image - One Value */
        oIoV: function oneImageOneValue(value, imageData, effect) {
            const module = Module;
            const ptr_ = allocateMemory(module, imageData.data.length);
            setMemory(module, imageData.data, ptr_);

            module[effect](ptr_, imageData.height, imageData.width, value);
            const ptr = new Uint8ClampedArray(module.HEAP8.buffer, ptr_, imageData.data.length);

            imageData.data.set(ptr);
            freeMemory(module, ptr_);
        },
        /* --------------------------------------------------------- */
        brightness: function(level, data) {
            this.oIoV(level, data, '_change_brightness');
        },
        saturation: function(level, data) {
            this.oIoV(level, data, '_change_saturation');
        },
        hue: function(level, data) {
            this.oIoV(level, data, '_change_hue');
        },
        transparency: function(level, data) {
            this.oIoV(level, data, '_change_transparency');
        },
        lighting: function(level, data) {
            this.oIoV(level, data, '_change_lighting');
        },
        intension_red: function(level, data) {
            this.oIoV(level, data, '_change_intension_red');
        },
        intension_green: function(level, data) {
            this.oIoV(level, data, '_change_intension_green');
        },
        intension_blue: function(level, data) {
            this.oIoV(level, data, '_change_intension_blue');
        },
        shades_of_gray: function(level, data) {
            this.oIoV(level, data, '_shades_of_gray');
        },
        sepia: function(level, data) {
            this.oIoV(level, data, '_sepia');
        },
        /* все эффекты, которые могут быть, будут перечислены тут */
    };

    /**
     * @param {Message} message 
     */
    function ApplyEffect(message) {
        if (isModuleLoaded) {
            message.effects.forEach(function(item) {
                effect[item.type](item.level, message.data);
            });
        } else {
            throw new Error('Module is not loaded!');
        }
    };
    window.onmessage = function(e) {
        ApplyEffect(e.data);
        postMessage(e.data.data);
    };
})(self);

