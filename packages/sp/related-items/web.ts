// import { _Web } from "../webs/types.js";
// import { RelatedItemManager, IRelatedItemManager } from "./types.js";

// declare module "../webs/types" {
//     interface _Web {
//         relatedItems: IRelatedItemManager;
//     }
//     interface IWeb {
//         /**
//          * The related items manager associated with this web
//          */
//         relatedItems: IRelatedItemManager;
//     }
// }

// Reflect.defineProperty(_Web.prototype, "relatedItems", {
//     configurable: true,
//     enumerable: true,
//     get: function (this: _Web) {
//         return RelatedItemManager(this);
//     },
// });
