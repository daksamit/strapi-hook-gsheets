"use strict";

const { GoogleSpreadsheet } = require("google-spreadsheet");

// interface SpreadsheetCredentials {
//   spreadsheet_id: string;
//   client_email: string;
//   private_key: string;
// }

// interface Spreadsheet {
//   doc: unknown;
//   getSheetByName: (string) => number;
// }

/**
 * Gsheets Hook
 */

module.exports = (strapi) => {
  return {
    /**
     * Default options
     * This object is merged to strapi.config.hook.settings.algolia
     */
    defaults: {
      spreadsheet_id: "",
      client_email: "",
      private_key: "",
    },

    /**
     * Initialize the hook
     */
    async initialize() {
      // Merging defaults en config/hook.json
      const { spreadsheet_id, client_email, private_key } = {
        ...this.defaults,
        ...strapi.config.hook.settings.gsheets,
      };

      if (
        !spreadsheet_id.length &&
        !client_email.length &&
        !private_key.length
      ) {
        throw Error("Gsheets Hook: Could not initialize: missing credentials");
      }

      const doc = new GoogleSpreadsheet(spreadsheet_id);

      await doc.useServiceAccountAuth({
        client_email,
        private_key,
      });
      await doc.loadInfo();

      const sheetsMap = new Map();
      const getSheets = () => {
        sheetsMap.clear();
        const sheets = Object.values(doc._rawSheets);
        sheets.forEach(
          (sheet: { _rawProperties: { title: string, index: number } }) =>
            sheetsMap.set(
              sheet._rawProperties.title,
              sheet._rawProperties.index
            )
        );
      };
      const getSheetByName = (name: string): number => {
        if (!sheetsMap.has(name)) {
          getSheets();
        }
        const index = sheetsMap.get(name);
        return doc.sheetsByIndex[index];
      };

      strapi.services.gsheets = {
        doc,
        getSheetByName,
      };
    },
  };
};
