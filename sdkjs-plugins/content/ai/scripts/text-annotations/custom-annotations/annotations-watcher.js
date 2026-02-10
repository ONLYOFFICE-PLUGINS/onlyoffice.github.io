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

/// <reference path="./types.js" />

class AnnotationsWatcher {
    constructor() {
        this._assistants = new Map();
        this._panel = null;
        /** @type {Map<string, AnnotationInfo>} */
        this._annotations = new Map();
    }

    showPanel() {
        if (this._panel) {
            this._closePanel();
        }
        const description = "List of annotations";

        let variation = {
            url: "customAnnotations.html",
            description: window.Asc.plugin.tr(description),
            isVisual: true,
            buttons: [],
            isModal: false,
            isCanDocked: false,
            type: "panel",
            EditorsSupport: ["word"],
        };

        this._panel = new window.Asc.PluginWindow();
        this._panel.attachEvent("onWindowReady", () => {
            this._annotations.forEach((annotationInfo) => {
                this._panel.command("onAddAnnotations", annotationInfo);
            });
        });
        this._panel.attachEvent("onAcceptAnnotation", (/** @type {RangeAddress} */ ctx) => {
            const assistant = this._assistants.get(ctx.assistantId);
            assistant.onAccept(ctx.paragraphId, ctx.rangeId);
        });
        this._panel.attachEvent("onRejectAnnotation", (/** @type {RangeAddress} */ ctx) => {
            const assistant = this._assistants.get(ctx.assistantId);
            assistant.onReject(ctx.paragraphId, ctx.rangeId);
        });

        this._panel.show(variation);
    }
    /** @param {Assistant} assistant */
    addAssistant(assistant) {
        if (this._assistants.has(assistant.assistantData.id)) {
            // console.warn("Assistant already added: " + assistant.assistantData.id);
        }
        this._assistants.set(assistant.assistantData.id, assistant);
        assistant.onRemoveAnnotation = this._onRemoveAnnotation.bind(this);
        assistant.onAddAnnotation = this._onAddAnnotation.bind(this);
    }

    /** @param {string} assistantId */
    removeAssistant(assistantId) {
        this._assistants.delete(assistantId);
    }

    /** @param {string} assistantId */
    hasAssistant(assistantId) {
        return this._assistants.has(assistantId);
    }

    _closePanel() {
        if (this._panel) {
            this._panel.close();
            this._panel = null;
        }
    }

    /** @param {AnnotationInfo} annotationInfo */
    _onAddAnnotation(annotationInfo) {
        const key = annotationInfo.paraId + "_" + annotationInfo.assistantId;
        this._annotations.set(key, annotationInfo);
        if (this._panel) {
            this._panel.command("onAddAnnotations", annotationInfo);
        }
    }

    /** @param {TextAnnotation | Array<TextAnnotation>} rangesInfo */
    _onRemoveAnnotation(rangesInfo) {
        if (!Array.isArray(rangesInfo)) {
            rangesInfo = [rangesInfo];
        }
        rangesInfo.forEach((rangeInfo) => {
            if (rangeInfo.name.slice(0, 16) !== "customAssistant_") {
                console.error("Remove annotation: not a custom annotation");
                return;
            }
            const key = rangeInfo.paragraphId + "_" + rangeInfo.name.slice(16);
            if (rangeInfo.all) {
                this._annotations.delete(key);
            } else {
                let annotation = this._annotations.get(key);
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
                    this._annotations.delete(key);
                }
            }
        });
        if (this._panel) {
            this._panel.command("onRemoveAnnotation", rangesInfo);
        }
    }
}
