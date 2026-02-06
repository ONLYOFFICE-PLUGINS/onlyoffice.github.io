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
    class AnnotationsPanel {
        constructor() {
            /** @type {Map<string, AnnotationInfo>} */
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

        /** @param {AnnotationInfo} annotationInfo */
        onAddAnnotations(annotationInfo) {
            const key = `${annotationInfo.paraId}_${annotationInfo.assistantId}`;
            this.annotations.set(key, annotationInfo);

            this._render();
        }

        /** @param {TextAnnotation | Array<TextAnnotation>} rangesInfo */
        onRemoveAnnotation(rangesInfo) {
            if (!Array.isArray(rangesInfo)) {
                rangesInfo = [rangesInfo];
            }
            rangesInfo.forEach((rangeInfo) => {
                if (rangeInfo.name.slice(0, 16) !== "customAssistant_") {
                    console.warn("Remove annotation: not a custom annotation");
                    return;
                }
                const key = rangeInfo.paragraphId + "_" + rangeInfo.name.slice(16);
                if (rangeInfo.all) {
                    this.annotations.delete(key);
                } else {
                    let annotation = this.annotations.get(key);
                    if (!annotation) {
                        return;
                    }

                    annotation.balloons.find((item, index) => {
                        if (item && item.rangeId === Number(rangeInfo.rangeId)) {
                            annotation.balloons[index] = null;
                            return true;
                        }
                        return false;
                    });

                    if (annotation.balloons.every(item => item === null)) {
                        this.annotations.delete(key);
                    }
                }
            });
            this._render();
        }

        _render() {
            const listContainer = this.listContainer;
            if (!listContainer) {
                return;
            }

            listContainer.innerHTML = "";

            const allAnnotations = Array.from(this.annotations.values());

            /** @type {Map<string, Map<string, Array<{annotationInfo: AnnotationInfo, item: AnnotationBalloonInfo}>>>} */
            const grouped = new Map();

            allAnnotations.forEach((annotationInfo) => {
                const assistantId = annotationInfo.assistantId;

                const paraId = annotationInfo.paraId || "";
                const balloons = annotationInfo.balloons;
                if (!balloons || !Array.isArray(balloons)) {
                    return;
                }

                balloons.forEach((item) => {
                    if (!item) {
                        return;
                    }

                    let byAssistant = grouped.get(assistantId);
                    if (!byAssistant) {
                        byAssistant = new Map();
                        grouped.set(assistantId, byAssistant);
                    }

                    let byPara = byAssistant.get(paraId);
                    if (!byPara) {
                        byPara = [];
                        byAssistant.set(paraId, byPara);
                    }

                    byPara.push({ annotationInfo, item });
                });
            });

            grouped.forEach((byPara, assistantId) => {
                listContainer.appendChild(this._createAssistantGroup(assistantId, byPara));
            });
        }

        /**
         * @param {string} assistantId
         * @param {Map<string, Array<{annotationInfo: AnnotationInfo, item: AnnotationBalloonInfo}>>} byPara
         */
        _createAssistantGroup(assistantId, byPara) {
            const group = document.createElement("div");
            group.className = "custom_annotations_group";

            const header = document.createElement("div");
            header.className = "custom_annotations_group_header";
            header.textContent = assistantId;

            group.appendChild(header);

            byPara.forEach((items, paraId) => {
                group.appendChild(this._createParagraphGroup(paraId, items));
            });

            return group;
        }

        /**
         * @param {string} paraId
         * @param {Array<{annotationInfo: AnnotationInfo, item: AnnotationBalloonInfo}>} items
         */
        _createParagraphGroup(paraId, items) {
            const group = document.createElement("div");
            group.className = "custom_annotations_subgroup";

            const header = document.createElement("div");
            header.className = "custom_annotations_subgroup_header";
            header.textContent = "Paragraph id: " + paraId;

            group.appendChild(header);

            items.forEach((item) => {
                group.appendChild(
                    this._createRangeItem(item.annotationInfo, item.item),
                );
            });

            return group;
        }

        /**
         * @param {AnnotationInfo} annotationInfo
         * @param {AnnotationBalloonInfo} item
         */
        _createRangeItem(annotationInfo, item) {
            const root = document.createElement("div");
            root.className = "custom_annotation_item";

            const assistantId = annotationInfo.assistantId;
            const name = `customAssistant_${assistantId}`;
            const paragraphId = annotationInfo.paraId;
            const rangeId = item.rangeId;
            
            /** @type {ReplaceInfoForPopup | HintInfoForPopup | ReplaceHintInfoForPopup} */
            let balloon = item.balloon;

            /** @type {RangeAddress} */
            const actionContext = {
                paragraphId,
                rangeId,
                assistantId,
            };

            root.addEventListener("click", () => {
                this._selectAnnotationInDocument(
                    paragraphId,
                    rangeId,
                    name,
                );
            });

            const header = document.createElement("div");
            header.className = "custom_annotation_header";

            const actions = document.createElement("div");
            actions.className = "custom_annotation_actions";

            const acceptBtn = document.createElement("button");
            acceptBtn.type = "button";
            acceptBtn.className = "custom_annotation_action custom_annotation_action_accept";
            acceptBtn.textContent = "✓";
            acceptBtn.addEventListener("click", (e) => {
                e.stopPropagation();
                this._onAcceptAnnotation(actionContext);
            });
            actions.appendChild(acceptBtn);

            if (balloon.type !== 0) { // Hint
                const rejectBtn = document.createElement("button");
                rejectBtn.type = "button";
                rejectBtn.className = "custom_annotation_action custom_annotation_action_reject";
                rejectBtn.textContent = "✕";
                rejectBtn.addEventListener("click", (e) => {
                    e.stopPropagation();
                    this._onRejectAnnotation(actionContext);
                });
                actions.appendChild(rejectBtn);
            }
            

            const meta = document.createElement("div");
            meta.className = "custom_annotation_meta";
            meta.textContent = balloon.original;

            const tooltip = this._createBalloonTooltip(balloon);

            header.appendChild(meta);
            header.appendChild(actions);
            root.appendChild(header);

            if (tooltip) {
                root.appendChild(tooltip);
            }

            return root;
        }

        /**
         * @param {ReplaceInfoForPopup | HintInfoForPopup | ReplaceHintInfoForPopup} balloon
         */
        _createBalloonTooltip(balloon) {
            /** @type {any} */
            const b = balloon;

            const hasSuggested = typeof b.suggested === "string" && b.suggested !== "";
            const hasExplanation = typeof b.explanation === "string" && b.explanation !== "";

            if (!hasSuggested && !hasExplanation) {
                return null;
            }

            const tooltip = document.createElement("div");
            tooltip.className = "custom_annotation_tooltip";

            if (hasSuggested) {
                const block = document.createElement("div");
                block.className = "custom_annotation_tooltip_block";

                const title = document.createElement("div");
                title.className = "custom_annotation_tooltip_title";
                title.textContent = window.Asc.plugin.tr("Suggested correction");

                const content = document.createElement("div");
                content.className = "custom_annotation_tooltip_content";

                const row = document.createElement("div");
                row.className = "custom_annotation_tooltip_row";

                const original = document.createElement("span");
                original.className = "custom_annotation_tooltip_original";
                original.textContent = balloon.original || "";

                const arrow = document.createElement("span");
                arrow.className = "custom_annotation_tooltip_arrow";
                arrow.textContent = "→";

                const corrected = document.createElement("span");
                corrected.className = "custom_annotation_tooltip_corrected";
                corrected.textContent = b.suggested;

                row.appendChild(original);
                row.appendChild(arrow);
                row.appendChild(corrected);

                content.appendChild(row);
                block.appendChild(title);
                block.appendChild(content);
                tooltip.appendChild(block);
            }

            if (hasExplanation) {
                const block = document.createElement("div");
                block.className = "custom_annotation_tooltip_block";

                const title = document.createElement("div");
                title.className = "custom_annotation_tooltip_title";
                title.textContent = window.Asc.plugin.tr("Explanation");

                const content = document.createElement("div");
                content.className = "custom_annotation_tooltip_content";
                content.textContent = b.explanation;

                block.appendChild(title);
                block.appendChild(content);
                tooltip.appendChild(block);
            }

            return tooltip;
        }

        /** @param {any} theme */
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

        /**
         * @param {string} paragraphId
         * @param {number} rangeId
         * @param {string} name
         */
        async _selectAnnotationInDocument(paragraphId, rangeId, name) {
            const annotation = {
                paragraphId: paragraphId,
                rangeId: rangeId,
                name: name,
            };

            await new Promise((resolve) => {
                window.Asc.plugin.executeMethod(
                    "SelectAnnotationRange",
                    [annotation],
                    resolve,
                );
            });
        }

        /**
         * @param {RangeAddress} ctx
         */
        _onAcceptAnnotation(ctx) {
            window.Asc.plugin.sendToPlugin("onAcceptAnnotation", ctx);
        }

        /**
         * @param {RangeAddress} ctx
         */
        _onRejectAnnotation(ctx) {
            window.Asc.plugin.sendToPlugin("onRejectAnnotation", ctx);
        }

    }

    const annotationsPanel = new AnnotationsPanel();
    window.Asc.plugin.init = annotationsPanel.init.bind(annotationsPanel);
    window.Asc.plugin.attachEvent("onAddAnnotations", annotationsPanel.onAddAnnotations.bind(annotationsPanel));
    window.Asc.plugin.attachEvent("onRemoveAnnotation", annotationsPanel.onRemoveAnnotation.bind(annotationsPanel));
    window.Asc.plugin.onTranslate = annotationsPanel.onTranslate.bind(annotationsPanel);
    window.Asc.plugin.onThemeChanged = annotationsPanel.onThemeChanged.bind(annotationsPanel);
    window.Asc.plugin.attachEvent("onThemeChanged", annotationsPanel.onThemeChanged.bind(annotationsPanel));


})(window);
