/*
 * (c) Copyright Ascensio System SIA 2010-2025
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-6 Ernesta Birznieka-Upish
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 */

// @ts-check

/// <reference path="./utils/theme.js" />
/// <reference path="./text-annotations/custom-annotations/types.js" />

(function (window) {
    const LOCAL_STORAGE_KEY = "onlyoffice_ai_saved_assistants";
    
    class AnnotationsPanel {
        constructor() {
            this.annotations = new Map();
            /** @type {HTMLElement | null} */
            this.container = document.getElementById("custom_annotations_panel");
            /** @type {HTMLElement | null} */
            this.listContainer = document.getElementById("custom_annotations_list");

            if (!this.container || !this.listContainer) {
                console.error("Custom Annotations: required elements are missing");
            }
        }

        init() {
            window.Asc.plugin.sendToPlugin("onWindowReady", {});

            if (!this.container) {
                return;
            }

            this.container.addEventListener("click", function (e) {
                if (e.target instanceof HTMLAnchorElement) {
                    e.preventDefault();
                    window.open(e.target.href, "_blank");
                }
            });
        }

        /** @param {EventAnnotationInfo} annotationInfo */
        onAddAnnotations(annotationInfo) {
            const key = `${annotationInfo.paraId}_${annotationInfo.assistantData.id}`;
            this.annotations.set(key, annotationInfo);

            this._render();
        }

        _render() {
            const listContainer = this.listContainer;
            if (!listContainer) {
                return;
            }

            listContainer.innerHTML = "";

            const allAnnotations = Array.from(this.annotations.values());
            console.log("Custom Annotations: all annotations", allAnnotations);

            /** @type {Map<string, Map<string, Array<{annotationInfo: EventAnnotationInfo, range: AnnotationRange, rangeIndex: number}>>>} */
            const grouped = new Map();

            allAnnotations.forEach((annotationInfo) => {
                const assistantName =
                    (annotationInfo.assistantData && annotationInfo.assistantData.name)
                        ? annotationInfo.assistantData.name
                        : (annotationInfo.assistantData && annotationInfo.assistantData.id)
                            ? annotationInfo.assistantData.id
                            : "Assistant";

                const paraId = annotationInfo.paraId || "";
                const ranges = annotationInfo.ranges;
                if (!ranges || !Array.isArray(ranges)) {
                    return;
                }

                ranges.forEach((range, rangeIndex) => {
                    if (!range) {
                        return;
                    }

                    let byAssistant = grouped.get(assistantName);
                    if (!byAssistant) {
                        byAssistant = new Map();
                        grouped.set(assistantName, byAssistant);
                    }

                    let byPara = byAssistant.get(paraId);
                    if (!byPara) {
                        byPara = [];
                        byAssistant.set(paraId, byPara);
                    }

                    byPara.push({ annotationInfo, range, rangeIndex });
                });
            });

            grouped.forEach((byPara, assistantName) => {
                listContainer.appendChild(this._createAssistantGroup(assistantName, byPara));
            });
        }

        /**
         * @param {string} assistantName
         * @param {Map<string, Array<{annotationInfo: EventAnnotationInfo, range: AnnotationRange, rangeIndex: number}>>} byPara
         */
        _createAssistantGroup(assistantName, byPara) {
            const group = document.createElement("div");
            group.className = "custom_annotations_group";

            const header = document.createElement("div");
            header.className = "custom_annotations_group_header";
            header.textContent = assistantName;

            group.appendChild(header);

            byPara.forEach((items, paraId) => {
                group.appendChild(this._createParagraphGroup(paraId, items));
            });

            return group;
        }

        /**
         * @param {string} paraId
         * @param {Array<{annotationInfo: EventAnnotationInfo, range: AnnotationRange, rangeIndex: number}>} items
         */
        _createParagraphGroup(paraId, items) {
            const group = document.createElement("div");
            group.className = "custom_annotations_subgroup";

            const header = document.createElement("div");
            header.className = "custom_annotations_subgroup_header";
            header.textContent = paraId;

            group.appendChild(header);

            items.forEach((item) => {
                group.appendChild(
                    this._createRangeItem(item.annotationInfo, item.range, item.rangeIndex),
                );
            });

            return group;
        }

        /**
         * @param {EventAnnotationInfo} annotationInfo
         * @param {AnnotationRange} range
         * @param {number} rangeIndex
         */
        _createRangeItem(annotationInfo, range, rangeIndex) {
            const root = document.createElement("div");
            root.className = "custom_annotation_item";

            const header = document.createElement("div");
            header.className = "custom_annotation_header";

            const title = document.createElement("div");
            title.className = "custom_annotation_title";
            title.textContent = `Range #${rangeIndex + 1}`;

            const meta = document.createElement("div");
            meta.className = "custom_annotation_meta";
            meta.textContent = `start=${range.start}, length=${range.length}`;

            header.appendChild(title);
            header.appendChild(meta);

            const kv = document.createElement("div");
            kv.className = "custom_annotation_kv";

            kv.appendChild(this._createKvRow("rangeId", range.id));
            kv.appendChild(this._createKvRow("recalcId", annotationInfo.recalcId));
            kv.appendChild(
                this._createKvRow(
                    "substring",
                    this._getRangeText(annotationInfo.text, range),
                ),
            );

            root.appendChild(header);
            root.appendChild(kv);

            return root;
        }

        /**
         * @param {string} text
         * @param {AnnotationRange} range
         */
        _getRangeText(text, range) {
            if (typeof text !== "string") {
                return "";
            }
            const start = typeof range.start === "number" ? range.start : 0;
            const length = typeof range.length === "number" ? range.length : 0;
            const end = start + Math.max(length, 0);
            return text.slice(start, end);
        }

        /**
         * @param {string} key
         * @param {any} value
         */
        _createKvRow(key, value) {
            const frag = document.createDocumentFragment();

            const k = document.createElement("div");
            k.className = "custom_annotation_k";
            k.textContent = key;

            const v = document.createElement("div");
            v.className = "custom_annotation_v";
            v.textContent = value === undefined || value === null ? "" : String(value);

            frag.appendChild(k);
            frag.appendChild(v);
            return frag;
        }

        /** @param {Array<AnnotationRange | null> | undefined} ranges */
        _formatRanges(ranges) {
            if (!ranges || !Array.isArray(ranges) || ranges.length === 0) {
                return "";
            }

            /** @type {AnnotationRange[]} */
            const safeRanges = ranges.filter((r) => r !== null);

            return safeRanges
                .map(r => `start=${r.start}, length=${r.length}, id=${r.id}`)
                .join("\n");
        }

        /** @param {any} obj */
        _safeStringify(obj) {
            try {
                return JSON.stringify(obj, null, 2);
            } catch (e) {
                return String(obj);
            }
        }

        onThemeChanged(theme) {
            window.Asc.plugin.onThemeChangedBase(theme);
            updateBodyThemeClasses(theme.type, theme.name);
            updateThemeVariables(theme);
        }

        onTranslate() {
            let elements = document.querySelectorAll(".i18n");
            elements.forEach(function (element) {
                if (
                    element instanceof HTMLTextAreaElement ||
                    element instanceof HTMLInputElement
                ) {
                    element.placeholder = window.Asc.plugin.tr(element.placeholder);
                } else if (element instanceof HTMLElement) {
                    element.innerText = window.Asc.plugin.tr(element.innerText);
                }
            });
        };
    }

    const annotationsPanel = new AnnotationsPanel();
    window.Asc.plugin.init = annotationsPanel.init.bind(annotationsPanel);
    window.Asc.plugin.attachEvent("onAddAnnotations", annotationsPanel.onAddAnnotations.bind(annotationsPanel));
    window.Asc.plugin.onTranslate = annotationsPanel.onTranslate.bind(annotationsPanel);
    window.Asc.plugin.onThemeChanged = annotationsPanel.onThemeChanged.bind(annotationsPanel);
    window.Asc.plugin.attachEvent("onThemeChanged", annotationsPanel.onThemeChanged.bind(annotationsPanel));


})(window);
