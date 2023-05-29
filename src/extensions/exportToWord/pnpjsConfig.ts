import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
/*import { AssignFrom } from "@pnp/core";
import Constants from "../../Constants/Constants";*/


// eslint-disable-next-line no-var
var _sp: SPFI = null;
let _graph: GraphFI = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (!!context) { // eslint-disable-line eqeqeq
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _sp;
};

export const getGraph = (context?: WebPartContext): GraphFI => {
  if (context !== null) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _graph;
};

/*export const getHelpDeskDemoSP = (): SPFI => {
  const sp = getSP()
  const spSite = spfi(Constants.HelpDeskDemoUrl).using(AssignFrom(sp.web));
  return spSite;
}*/
