import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI | null = null; // ✅ ALLOW null
export const setupPnP = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};

export const getSP = (): SPFI => {
  if (_sp === null) {
    throw new Error("PnP JS is not initialized. Call setupPnP(context) first.");
  }
  return _sp;
};
