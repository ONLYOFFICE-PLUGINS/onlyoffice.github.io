/**
 * @typedef {Object} localStorageCustomAssistantItem
 * @property {string} id
 * @property {string} name
 * @property {number} type
 * @property {string} query
 */

/**
 * @typedef {Object} HintAiResponse
 * @property {string} origin
 * @property {string} reason
 * @property {number} paragraph
 * @property {number} occurrence
 * @property {number} confidence
 */

/**
 * @typedef {Object} HintInfoForPopup
 * @property {string} original
 * @property {string} explanation
 * @property {number} type
 */

/**
 * @typedef {Object} ReplaceAiResponse
 * @property {string} origin
 * @property {string} suggestion
 * @property {string} reason
 * @property {number} paragraph
 * @property {number} occurrence
 * @property {number} confidence
 */

/**
 * @typedef {Object} ReplaceInfoForPopup
 * @property {string} original
 * @property {string} suggested
 * @property {number} type
 */

/**
 * @typedef {Object} ReplaceHintAiResponse
 * @property {string} origin
 * @property {string} suggestion
 * @property {string} reason
 * @property {string} difference
 * @property {number} paragraph
 * @property {number} occurrence
 * @property {number} confidence
 */

/**
 * @typedef {Object} ReplaceHintInfoForPopup
 * @property {string} original
 * @property {string} suggested
 * @property {string} explanation
 * @property {number} type
 */

/**
 * @typedef {ReplaceHintInfoForPopup | HintInfoForPopup | ReplaceInfoForPopup} InfoForPopup
 */

/**
 * @typedef {CustomAnnotator & TextAnnotator & AssistantHint & AssistantReplaceHint & AssistantReplace} Assistant
 */

/**
 * @typedef {Object} AnnotationRange
 * @property {number} start
 * @property {number} length
 * @property {number} id
 */

/**
 * @typedef {Object} AnnotationBalloonInfo
 * @property {number} rangeId
 * @property {ReplaceInfoForPopup | HintInfoForPopup | ReplaceHintInfoForPopup} balloon
 */

/**
 * @typedef {Object} AnnotationInfo 
 * @property {string} paraId
 * @property {string} recalcId
 * @property {Array<AnnotationBalloonInfo | null>} balloons
 * @property {string} assistantId
*/

/**
 * @typedef {Object} TextAnnotation 
 * @property {string} name
 * @property {string} rangeId
 * @property {string} paragraphId
 * @property {boolean} [all]
*/

/** 
 * @typedef {Object} RangeAddress
 * @property {string} paragraphId
 * @property {number} rangeId
 * @property {string} assistantId
 */