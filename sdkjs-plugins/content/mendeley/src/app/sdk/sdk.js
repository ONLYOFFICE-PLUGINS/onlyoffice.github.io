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

// @ts-check

/// <reference path="../types-global.js" />
/// <reference path="./types.js" />

import { MendeleyToCls } from "./mendeley-to-csl";

class Sdk {
    /** @param {{authFlow: any}} authFlow */
    constructor(authFlow) {
        this._mendeleySdk = MendeleySDK(authFlow);
        this._userId = 0;
        /** @type {Array<UserGroupInfo>}} */
        this._userGroups = [];
    }

    /**
     * Get items from user library
     * @param {string|null} search
     * @param {string[]} [itemsID]
     * @param {string} [format]
     * @returns {Promise<SearchResult>}
     */
    getItems(search, itemsID, format) {
        let promise = Promise.resolve({items: []});

         /*this._mendeleySdk.documents.list({
                limit: 6,
                view: "all",
            }).then((response) => {
                console.error(response);
            });
         */
        if (search) {
            promise = this._mendeleySdk.documents.search({
                query: search,
                limit: 20,
                view: "bib",
            });
        } else if (itemsID || format) {
            // In mendeley sdk this way doesn't work (But for zotero it does)
        } else {
            promise = this._mendeleySdk.documents.list({
                limit: 16,
                view: "bib",
                sort: "last_modified",
                order: "desc"
            });
        }
        return promise.then((response) => {
            console.warn(response.items);
            response.items.forEach(MendeleyToCls.transform.bind(MendeleyToCls));
            return response;
        });
    }

    /**
     * Get items from group library
     * @param {string | null} search
     * @param {number|string} groupId
     * @param {string[]} [itemsID]
     * @returns {Promise<SearchResult>}
     */
    getGroupItems(search, groupId, itemsID) {
        var self = this;

        return new Promise(function (resolve, reject) {
            resolve({items: []});
        });
    }

    /**
     * Get user groups
     * @returns {Promise<Array<UserGroupInfo>>}
     */
    getUserGroups() {
        var self = this;

        this._mendeleySdk.folders.list({
                limit: 6
            }).then((response) => {
                console.error(response);
            });

        return new Promise(function (resolve, reject) {
            if (self._userGroups.length > 0) {
                resolve(self._userGroups);
                return;
            }

            resolve(self._userGroups);
        });
    }

}

export { Sdk };
