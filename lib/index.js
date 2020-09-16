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

//  NOTE: don't forget to share sheets with email service account

module.exports = (strapi) => {
  return {
    /**
     * Default options
     * This object is merged to strapi.config.hook.settings.gsheets
     */
    defaults: {
      spreadsheets: [],
      client_email: "",
      private_key: "",
    },

    /**
     * Initialize the hook
     */
    async initialize() {
      // Merging defaults en config/hook.json
      const { spreadsheets, client_email, private_key } = {
        ...this.defaults,
        ...strapi.config.hook.settings.gsheets,
      };

      strapi.hook.gsheets.spreadsheets = new Map(spreadsheets);
      const spreadsheetsMap = new Map(spreadsheets);

      strapi.services.gsheets = new Map();

      for (let [sheetName, options] of spreadsheetsMap) {
        // if is spreadsheet_id
        if (typeof options === "string") {
          const doc = new GoogleSpreadsheet(options);
          if (!client_email || !private_key) {
            throw Error(
              `Gsheets Hook: Could not initialize: missing client_email or private_key for "${sheetName}"`
            );
          }
          await doc.useServiceAccountAuth({ client_email, private_key });
          await doc.loadInfo();

          strapi.services.gsheets.set(sheetName, doc);
          //
        } else {
          if (!options.spreadsheet_id) {
            throw Error(
              `Gsheets Hook: Could not initialize: missing spreadsheet_id for "${sheetName}"`
            );
          }
          const doc = new GoogleSpreadsheet(options.spreadsheet_id);
          await doc.useServiceAccountAuth({
            client_email: options.client_email || client_email,
            private_key: options.private_key || private_key,
          });
          await doc.loadInfo();

          strapi.services.gsheets.set(sheetName, doc);
        }
      }
    },
  };
};
