/**
 *
 * (c) Copyright Ascensio System SIA 2020
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 */

(function(window, undefined) {
	var Ps1, Ps2;
	// var proxy = 'https://plugins-services.onlyoffice.com/proxy';
	// todo use your local proxy server for test this plugin
	var proxy = 'http://127.0.0.1:5000';
	var displayNoneClass = "display-none";
	var elements = {};

	function showLoader(bShow) {
		if (elements.loader)
		switchClass(elements.loader, displayNoneClass, !bShow);
	};

	function switchClass(el, className, add) {
		if (add) {
			el.classList.add(className);
		} else {
			el.classList.remove(className);
		}
	};

	window.oncontextmenu = function(e) {
		if (e.preventDefault)
			e.preventDefault();
		if (e.stopPropagation)
			e.stopPropagation();
		return false;
	};

	var sEmptyQuery = '',
		sLastQuery = sEmptyQuery,
		nImageWidth = 90,
		nVertGap = 15,
		nLastPage = 1,
		nLastPageCount = 1;
	
	function createScript(oElement, w, h) {
		var sScript = '';

		if (oElement) {
			switch (window.Asc.plugin.info.editorType) {
				case 'word': {
					sScript += 'var oDocument = Api.GetDocument();';
					sScript += '\nvar oParagraph, oRun, arrInsertResult = [], oImage;';
					sScript += '\noParagraph = Api.CreateParagraph();';
					sScript += '\narrInsertResult.push(oParagraph);';
					var sSrc = oElement.Src;
					var nEmuWidth = ((w / 96) * 914400) >> 0;
					var nEmuHeight = ((h / 96) * 914400) >> 0;
					sScript += '\n oImage = Api.CreateImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ');';
					sScript += '\noParagraph.AddDrawing(oImage);';
					sScript += '\noDocument.InsertContent(arrInsertResult);';
					break;
				}
				case 'slide':{
					sScript += 'var oPresentation = Api.GetPresentation();';
					sScript += '\nvar oSlide = oPresentation.GetCurrentSlide();';
					sScript += '\nif(oSlide) {';
					sScript += '\nvar fSlideWidth = oSlide.GetWidth(), fSlideHeight = oSlide.GetHeight();';
					var sSrc = oElement.Src;
					var nEmuWidth = ((w / 96) * 914400) >> 0;
					var nEmuHeight = ((h / 96) * 914400) >> 0;
					sScript += '\n var oImage = Api.CreateImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ');';
					sScript += '\n oImage.SetPosition((fSlideWidth -' + nEmuWidth +  ')/2, (fSlideHeight -' + nEmuHeight +  ')/2);';
					sScript += '\n oSlide.AddObject(oImage);';
					sScript += '\n}'
					break;
				}
				case 'cell':{
					sScript += '\nvar oWorksheet = Api.GetActiveSheet();';
					sScript += '\nif(oWorksheet) {';
					sScript += '\nvar oActiveCell = oWorksheet.GetActiveCell();';
					sScript += '\nvar nCol = oActiveCell.GetCol(), nRow = oActiveCell.GetRow();';
					var sSrc = oElement.Src;
					var nEmuWidth = ((w / 96) * 914400) >> 0;
					var nEmuHeight = ((h / 96) * 914400) >> 0;
					sScript += '\n oImage = oWorksheet.AddImage(\'' + sSrc + '\', ' + nEmuWidth + ', ' + nEmuHeight + ', nCol, 0, nRow, 0);';
					sScript += '\n}';
					break;
				}
			}
		}
		return sScript;
	};
	
	window.Asc.plugin.init = function() {
		elements = {
			loader: document.getElementById("loader-container"),
			contentHolder: document.getElementById("main-container-id"),
			container: document.getElementById('scrollable-container-id')
		};

		Ps1 = new PerfectScrollbar('#scrollable-container-id', {});
		Ps2 = new PerfectScrollbar('#main-container-id', {});


		$(window).resize(function() {
			updatePaddings();
			updateScroll();
			updateNavigation();
		});

		$('input').keydown(function(e) {
			if(e.keyCode === 13)
				$('#button-search-id').trigger('click');
		});

		$('#button-search-id').click(function() {
			sLastQuery = $('#search-id').val().trim().replace(/ /gi,'+');
			if(sLastQuery === '') {
				sLastQuery = sEmptyQuery;
			}
			loadClipArtPage(1, sLastQuery);
			return false;
		});

		$('#navigation-first-page-id').click(function(e) {
			if(nLastPage > 1) {
				loadClipArtPage(1, sLastQuery);
			}
		});
	
		$('#navigation-prev-page-id').click(function(e) {
			if(nLastPage > 1) {
				loadClipArtPage(Number(nLastPage) - 1, sLastQuery);
			}
		});
	
		$('#navigation-next-page-id').click(function(e) {
			if(nLastPage < nLastPageCount) {
				loadClipArtPage(Number(nLastPage) + 1, sLastQuery);
			}
		});
	
		$('#navigation-last-page-id').click(function(e) {
			if(nLastPage < nLastPageCount) {
				loadClipArtPage(Number(nLastPageCount), sLastQuery);
			}
		});

		// updateScroll();
		// loadClipArtPage(1, sLastQuery);
	};

	function CreateRequest() {
		var Request;

		if (window.XMLHttpRequest) {
			Request = new XMLHttpRequest();
		} else if (window.ActiveXObject) {
			//Internet explorer
			try{
				Request = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (CatchException) {
				Request = new ActiveXObject("Msxml2.XMLHTTP");
			}
		}

		if (!Request) {
			alert(window.Asc.plugin.tr("Unable to create the XMLHttpRequest"));
		}

		return Request;
	};

	function fillTableFromResponse(imgsInfo) {
		var oContainer = $('#preview-images-container-id');
		oContainer.empty();

		//calculate count images in string
		var nFullWidth = $('#scrollable-container-id').width() - 20;
		var nCount = (nFullWidth / (nImageWidth + 2 * nVertGap) + 0.01) >> 0;
		if(nCount < 1)
			nCount = 1;

		var nGap = 0;
		nGap = ( ( ( nFullWidth - nCount * nImageWidth ) / ( nCount ) ) / 2 ) >> 0;

		for (var i = 0; i < imgsInfo.length; ++i) {
			var oDivElement = $('<div></div>');
			oDivElement.css('display', 'inline-block');
			oDivElement.css('width', nImageWidth + 'px');
			oDivElement.css('height', nImageWidth + 'px');
			oDivElement.css('vertical-align','middle');
			$(oDivElement).addClass('noselect');
			oDivElement.css('margin-left', nGap + 'px');
			oDivElement.css('margin-right', nGap + 'px');
			oDivElement.css('margin-bottom', nVertGap + 'px');

			var oImageTh = {
				width : imgsInfo[i]["Width"],
				height : imgsInfo[i]["Height"]
			};
			var nMaxSize = Math.max(oImageTh.width, oImageTh.height);
			var fCoeff = nImageWidth/nMaxSize;
			var oImgElement = $('<img>');
			var nWidth = (oImageTh.width * fCoeff) >> 0;
			var nHeight = (oImageTh.height * fCoeff) >> 0;
			if (nWidth === 0 || nHeight === 0) {
				 oImgElement.on('load', function(event) {
					var nMaxSize = Math.max(this.naturalWidth, this.naturalHeight);
					var fCoeff = nImageWidth / nMaxSize;
					var nWidth = (this.naturalWidth * fCoeff) >> 0;
					var nHeight = (this.naturalHeight * fCoeff) >> 0;

					$(this).css('width', nWidth + 'px');
					$(this).css('height', nHeight + 'px');
					$(this).css('margin-left', (((nImageWidth - nWidth)/2) >> 0) + 'px');
					$(this).css('margin-top', (((nImageWidth - nHeight)/2) >> 0) + 'px');
				 });
			}
			oImgElement.css('width', nWidth + 'px');
			oImgElement.css('height', nHeight + 'px');
			oImgElement.css('margin-left', (((nImageWidth - nWidth)/2) >> 0) + 'px');
			oImgElement.css('margin-top', (((nImageWidth - nHeight)/2) >> 0) + 'px');
			oImgElement.attr('src',  imgsInfo[i].Src);
			oImgElement.attr('data-index', i + '');
			oImgElement.mouseenter(function(e) {
				$(this).css('opacity', '0.65');
			});
			oImgElement.mouseleave(function(e) {
				$(this).css('opacity', '1');
			});

			function addImg(img) {
				window.Asc.plugin.info.recalculate = true;
				var oElement = imgsInfo[parseInt(img.dataset.index)];
				window.Asc.plugin.executeCommand("command", createScript(oElement, img.naturalWidth, img.naturalHeight), function() {
					img.style.pointerEvents = "auto";
				});
			};

			oImgElement.click(function(e) {
				img.style.pointerEvents = "none";
				addImg(this);
			});

			oImgElement.on('dragstart', function(event) { event.preventDefault(); });

			oDivElement.append(oImgElement);
			oContainer.append(oDivElement);
		}
		updateScroll();
		showLoader(false);
	};

	function updateNavigation() {
		if(arguments.length == 2) {
			nLastPage = arguments[0];
			nLastPageCount = arguments[1];
		}

		if (nLastPage < nLastPageCount)
			var nUsePage = nLastPage - 1;

		var oPagesCell = $('#pages-cell-id');
		oPagesCell.empty();
		var nW = $('#pagination-table-container-id').width() - $('#pagination-table-id').width();
		var nMaxCountPages = (nW / 22) >> 0;

		if(nLastPageCount === 0) {
			$('#pagination-table-id').hide();
			return;
		} else {
			$('#pagination-table-id').show();
		}

		var nStart, nEnd;
		if(nLastPageCount <= nMaxCountPages) {
			nStart = 0;
			nEnd = nLastPageCount;
		} else if(nUsePage < nMaxCountPages) {
			nStart = 0;
			nEnd = nMaxCountPages;
		} else if( (nLastPageCount - nUsePage) <= nMaxCountPages) {
			nStart = nLastPageCount - nMaxCountPages;
			nEnd = nLastPageCount;
		} else {
			nStart = nUsePage - ( ( nMaxCountPages / 2 ) >> 0 );
			nEnd = nUsePage + ( ( nMaxCountPages / 2 ) >> 0 );
		}
		for(var i = nStart;  i < nEnd; ++i) {
			var oButtonElement = $('<div class="pagination-button-div noselect" style="width:22px; height:22px;"><p>' + (i + 1) +'</p></div>');
			oPagesCell.append(oButtonElement);
			oButtonElement.attr('data-index', i + '');
			if(i === nUsePage) {
				oButtonElement.addClass('pagination-button-div-selected');
			}
			oButtonElement.click(function(e) {
				$(this).addClass('pagination-button-div-selected');
				loadClipArtPage(parseInt($(this).attr('data-index')) + 1, sLastQuery);
			});
		}
	};

	function SendRequest(r_method, r_path, r_args, testData) {
		var Request = CreateRequest();

		if (!Request)
			return;

		showLoader(true);

		Request.onreadystatechange = function() {
			if (Request.readyState == 4) {
				if (Request.status == 200) {
					try {
						$('#preview-images-container-id').empty();
						var parser = new DOMParser();
						var doc = parser.parseFromString(JSON.parse(Request.responseText), "text/html");
						var docImgs = $('.artwork img', doc);
						var imgsInfo = [];
						var pagesInfo = $('.page-link', doc)[1].innerText.split(" / ");
						var current_page = pagesInfo[0];
						var allPages = pagesInfo[1];
						var imgCount = docImgs.length;
						elements.container.scrollTop = 0;
						Ps1.update();
	
						//setting correct url for each image
						docImgs.each(function() {
							$(this).attr("src", "https://openclipart.org" + $(this).attr("src"));
						});
						updateNavigation(Number(current_page), Number(allPages));
	
						if (imgCount === 0) {
							showLoader(false);
							$('#preview-images-container-id').empty();
							$('<div>', {
								"class": "no-results",
								text: window.Asc.plugin.tr("No results.")
							}).appendTo('#preview-images-container-id');
							updateScroll();
							return;
						}
	
						for (var imgIdx = 0; imgIdx < imgCount; imgIdx++) {
							var img = new Image();
							img.onload = function() {
								var imgInfo = {
								"Width": this.width,
								"Height": this.height,
								"Src": this.src,
								"HTML": this.outerHTML
								};
	
								if (imgInfo.width !== 0 && imgInfo.height !== 0)
									imgsInfo.push(imgInfo);
								else
									imgCount--;
	
								if (imgsInfo.length == imgCount)
									fillTableFromResponse(imgsInfo);
							};
							img.onerror = function() {
								imgCount--;
								if (imgsInfo.length == imgCount)
									fillTableFromResponse(imgsInfo);
							};
							img.src = $(docImgs[imgIdx]).attr('src');
						}
					} catch (error) {
						handleError();
					}
				} else {
					handleError();
				}
			}
		};

		if (r_method.toLowerCase() == "get" && r_args.length > 0)
		r_path += "?" + r_args;

		let settings = {
			target: r_path,
			method: 'GET',
			data: testData
		};

		Request.open(r_method, proxy, true);

		if (r_method.toLowerCase() == "post") {
			Request.send(JSON.stringify(settings));
		} else {
			Request.send(null);
		}
	};

	function handleError() {
		elements.container.scrollTop = 0;
		Ps1.update();
		updateNavigation(0, 0);
		var oContainer = $('#preview-images-container-id');
		oContainer.empty();
		var oParagraph = $('<p style=\"font-size: 15px; font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif;\">' + window.Asc.plugin.tr("Error has occured when loading data.") + '</p>');
		oContainer.append(oParagraph);
		showLoader(false);
	};

	function updatePaddings() {
		var oContainer = $('#preview-images-container-id');
		var nFullWidth = $('#scrollable-container-id').width() - 20;
		var nCount = (nFullWidth / (nImageWidth + 2 * nVertGap) + 0.01) >> 0;
		if (nCount < 1)
			nCount = 1;

		var nGap = ( ( ( nFullWidth - nCount * nImageWidth ) / ( nCount ) ) / 2 ) >> 0;
		var aChildNodes = oContainer[0].childNodes;

		for (var i = 0; i < aChildNodes.length; ++i) {
			var oDivElement = aChildNodes[i];
			$(oDivElement).css('margin-left', nGap + 'px');
			$(oDivElement).css('margin-right', nGap + 'px');
		}
	};

	function loadClipArtPage(nIndex, sQuery) {
		// SendRequest("POST", 'https://openclipart.org/search/?query=' + sQuery + '&p=' + nIndex, "");
		SendRequest("POST", 'https://openclipart.org/search/', null, {query: sQuery , p: nIndex});
	};

	function updateScroll() {
		Ps1.update();
		Ps2.update();
	};

	window.Asc.plugin.onTranslate = function() {
		var elements = document.getElementsByClassName("i18n");
		for (var i = 0; i < elements.length; i++) {
			var el = elements[i];
			if (el.attributes["placeholder"]) el.attributes["placeholder"].value = window.Asc.plugin.tr(el.attributes["placeholder"].value);
			if (el.innerText) el.innerText = window.Asc.plugin.tr(el.innerText);
		}
	};

	window.Asc.plugin.button = function(id) {
		this.executeCommand("close", '');
	};

	window.Asc.plugin.onExternalMouseUp = function() {
		var evt = document.createEvent("MouseEvents");
		evt.initMouseEvent("mouseup", true, true, window, 1, 0, 0, 0, 0,
			false, false, false, false, 0, null);

		document.dispatchEvent(evt);
	};
})(window);