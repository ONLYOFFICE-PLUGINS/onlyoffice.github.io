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


    //polyfill

    var Module=typeof Module!="undefined"?Module:{};var moduleOverrides=Object.assign({},Module);var arguments_=[];var thisProgram="./this.program";var quit_=(status,toThrow)=>{throw toThrow};var ENVIRONMENT_IS_WEB=typeof window=="object";var ENVIRONMENT_IS_WORKER=typeof importScripts=="function";var ENVIRONMENT_IS_NODE=typeof process=="object"&&typeof process.versions=="object"&&typeof process.versions.node=="string";var scriptDirectory="";function locateFile(path){if(Module["locateFile"]){return Module["locateFile"](path,scriptDirectory)}return scriptDirectory+path}var read_,readAsync,readBinary,setWindowTitle;if(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER){if(ENVIRONMENT_IS_WORKER){scriptDirectory=self.location.href}else if(typeof document!="undefined"&&document.currentScript){scriptDirectory=document.currentScript.src}if(scriptDirectory.indexOf("blob:")!==0){scriptDirectory=scriptDirectory.substr(0,scriptDirectory.replace(/[?#].*/,"").lastIndexOf("/")+1)}else{scriptDirectory=""}{read_=url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.send(null);return xhr.responseText};if(ENVIRONMENT_IS_WORKER){readBinary=url=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,false);xhr.responseType="arraybuffer";xhr.send(null);return new Uint8Array(xhr.response)}}readAsync=(url,onload,onerror)=>{var xhr=new XMLHttpRequest;xhr.open("GET",url,true);xhr.responseType="arraybuffer";xhr.onload=()=>{if(xhr.status==200||xhr.status==0&&xhr.response){onload(xhr.response);return}onerror()};xhr.onerror=onerror;xhr.send(null)}}setWindowTitle=title=>document.title=title}else{}var out=Module["print"]||console.log.bind(console);var err=Module["printErr"]||console.warn.bind(console);Object.assign(Module,moduleOverrides);moduleOverrides=null;if(Module["arguments"])arguments_=Module["arguments"];if(Module["thisProgram"])thisProgram=Module["thisProgram"];if(Module["quit"])quit_=Module["quit"];var wasmBinary;if(Module["wasmBinary"])wasmBinary=Module["wasmBinary"];var noExitRuntime=Module["noExitRuntime"]||true;if(typeof WebAssembly!="object"){abort("no native wasm support detected")}var wasmMemory;var ABORT=false;var EXITSTATUS;var buffer,HEAP8,HEAPU8,HEAP16,HEAPU16,HEAP32,HEAPU32,HEAPF32,HEAPF64;function updateGlobalBufferAndViews(buf){buffer=buf;Module["HEAP8"]=HEAP8=new Int8Array(buf);Module["HEAP16"]=HEAP16=new Int16Array(buf);Module["HEAP32"]=HEAP32=new Int32Array(buf);Module["HEAPU8"]=HEAPU8=new Uint8Array(buf);Module["HEAPU16"]=HEAPU16=new Uint16Array(buf);Module["HEAPU32"]=HEAPU32=new Uint32Array(buf);Module["HEAPF32"]=HEAPF32=new Float32Array(buf);Module["HEAPF64"]=HEAPF64=new Float64Array(buf)}var INITIAL_MEMORY=Module["INITIAL_MEMORY"]||134217728;var wasmTable;var __ATPRERUN__=[];var __ATINIT__=[];var __ATPOSTRUN__=[onLoadModule];var runtimeInitialized=false;function preRun(){if(Module["preRun"]){if(typeof Module["preRun"]=="function")Module["preRun"]=[Module["preRun"]];while(Module["preRun"].length){addOnPreRun(Module["preRun"].shift())}}callRuntimeCallbacks(__ATPRERUN__)}function initRuntime(){runtimeInitialized=true;callRuntimeCallbacks(__ATINIT__)}function postRun(){if(Module["postRun"]){if(typeof Module["postRun"]=="function")Module["postRun"]=[Module["postRun"]];while(Module["postRun"].length){addOnPostRun(Module["postRun"].shift())}}callRuntimeCallbacks(__ATPOSTRUN__)}function addOnPreRun(cb){__ATPRERUN__.unshift(cb)}function addOnInit(cb){__ATINIT__.unshift(cb)}function addOnPostRun(cb){__ATPOSTRUN__.unshift(cb)}var runDependencies=0;var runDependencyWatcher=null;var dependenciesFulfilled=null;function addRunDependency(id){runDependencies++;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}}function removeRunDependency(id){runDependencies--;if(Module["monitorRunDependencies"]){Module["monitorRunDependencies"](runDependencies)}if(runDependencies==0){if(runDependencyWatcher!==null){clearInterval(runDependencyWatcher);runDependencyWatcher=null}if(dependenciesFulfilled){var callback=dependenciesFulfilled;dependenciesFulfilled=null;callback()}}}function abort(what){{if(Module["onAbort"]){Module["onAbort"](what)}}what="Aborted("+what+")";err(what);ABORT=true;EXITSTATUS=1;what+=". Build with -sASSERTIONS for more info.";var e=new WebAssembly.RuntimeError(what);throw e}var dataURIPrefix="data:application/octet-stream;base64,";function isDataURI(filename){return filename.startsWith(dataURIPrefix)}var wasmBinaryFile;wasmBinaryFile="effects.wasm";if(!isDataURI(wasmBinaryFile)){wasmBinaryFile=locateFile(wasmBinaryFile)}function getBinary(file){try{if(file==wasmBinaryFile&&wasmBinary){return new Uint8Array(wasmBinary)}if(readBinary){return readBinary(file)}else{throw"both async and sync fetching of the wasm failed"}}catch(err){abort(err)}}function getBinaryPromise(){if(!wasmBinary&&(ENVIRONMENT_IS_WEB||ENVIRONMENT_IS_WORKER)){if(typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){if(!response["ok"]){throw"failed to load wasm binary file at '"+wasmBinaryFile+"'"}return response["arrayBuffer"]()}).catch(function(){return getBinary(wasmBinaryFile)})}}return Promise.resolve().then(function(){return getBinary(wasmBinaryFile)})}function createWasm(){var info={"a":asmLibraryArg};function receiveInstance(instance,module){var exports=instance.exports;Module["asm"]=exports;wasmMemory=Module["asm"]["c"];updateGlobalBufferAndViews(wasmMemory.buffer);wasmTable=Module["asm"]["e"];addOnInit(Module["asm"]["d"]);removeRunDependency("wasm-instantiate")}addRunDependency("wasm-instantiate");function receiveInstantiationResult(result){receiveInstance(result["instance"])}function instantiateArrayBuffer(receiver){return getBinaryPromise().then(function(binary){return WebAssembly.instantiate(binary,info)}).then(function(instance){return instance}).then(receiver,function(reason){err("failed to asynchronously prepare wasm: "+reason);abort(reason)})}function instantiateAsync(){if(!wasmBinary&&typeof WebAssembly.instantiateStreaming=="function"&&!isDataURI(wasmBinaryFile)&&typeof fetch=="function"){return fetch(wasmBinaryFile,{credentials:"same-origin"}).then(function(response){var result=WebAssembly.instantiateStreaming(response,info);return result.then(receiveInstantiationResult,function(reason){err("wasm streaming compile failed: "+reason);err("falling back to ArrayBuffer instantiation");return instantiateArrayBuffer(receiveInstantiationResult)})})}else{return instantiateArrayBuffer(receiveInstantiationResult)}}if(Module["instantiateWasm"]){try{var exports=Module["instantiateWasm"](info,receiveInstance);return exports}catch(e){err("Module.instantiateWasm callback failed with error: "+e);return false}}instantiateAsync();return{}}function callRuntimeCallbacks(callbacks){while(callbacks.length>0){callbacks.shift()(Module)}}function _abort(){abort("")}function abortOnCannotGrowMemory(requestedSize){abort("OOM")}function _emscripten_resize_heap(requestedSize){var oldSize=HEAPU8.length;requestedSize=requestedSize>>>0;abortOnCannotGrowMemory(requestedSize)}var asmLibraryArg={"b":_abort,"a":_emscripten_resize_heap};var asm=createWasm();var ___wasm_call_ctors=Module["___wasm_call_ctors"]=function(){return(___wasm_call_ctors=Module["___wasm_call_ctors"]=Module["asm"]["d"]).apply(null,arguments)};var _change_brightness=Module["_change_brightness"]=function(){return(_change_brightness=Module["_change_brightness"]=Module["asm"]["f"]).apply(null,arguments)};var _change_saturation=Module["_change_saturation"]=function(){return(_change_saturation=Module["_change_saturation"]=Module["asm"]["g"]).apply(null,arguments)};var _change_transparency=Module["_change_transparency"]=function(){return(_change_transparency=Module["_change_transparency"]=Module["asm"]["h"]).apply(null,arguments)};var _change_hue=Module["_change_hue"]=function(){return(_change_hue=Module["_change_hue"]=Module["asm"]["i"]).apply(null,arguments)};var _change_lighting=Module["_change_lighting"]=function(){return(_change_lighting=Module["_change_lighting"]=Module["asm"]["j"]).apply(null,arguments)};var _change_intension_red=Module["_change_intension_red"]=function(){return(_change_intension_red=Module["_change_intension_red"]=Module["asm"]["k"]).apply(null,arguments)};var _change_intension_green=Module["_change_intension_green"]=function(){return(_change_intension_green=Module["_change_intension_green"]=Module["asm"]["l"]).apply(null,arguments)};var _change_intension_blue=Module["_change_intension_blue"]=function(){return(_change_intension_blue=Module["_change_intension_blue"]=Module["asm"]["m"]).apply(null,arguments)};var _shades_of_gray=Module["_shades_of_gray"]=function(){return(_shades_of_gray=Module["_shades_of_gray"]=Module["asm"]["n"]).apply(null,arguments)};var _sepia=Module["_sepia"]=function(){return(_sepia=Module["_sepia"]=Module["asm"]["o"]).apply(null,arguments)};var _malloc=Module["_malloc"]=function(){return(_malloc=Module["_malloc"]=Module["asm"]["p"]).apply(null,arguments)};var _free=Module["_free"]=function(){return(_free=Module["_free"]=Module["asm"]["q"]).apply(null,arguments)};var calledRun;dependenciesFulfilled=function runCaller(){if(!calledRun)run();if(!calledRun)dependenciesFulfilled=runCaller};function run(args){args=args||arguments_;if(runDependencies>0){return}preRun();if(runDependencies>0){return}function doRun(){if(calledRun)return;calledRun=true;Module["calledRun"]=true;if(ABORT)return;initRuntime();if(Module["onRuntimeInitialized"])Module["onRuntimeInitialized"]();postRun()}if(Module["setStatus"]){Module["setStatus"]("Running...");setTimeout(function(){setTimeout(function(){Module["setStatus"]("")},1);doRun()},1)}else{doRun()}}if(Module["preInit"]){if(typeof Module["preInit"]=="function")Module["preInit"]=[Module["preInit"]];while(Module["preInit"].length>0){Module["preInit"].pop()()}}run();

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

