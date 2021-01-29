import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface ISpfxFluentuiPivotProps {
  site: IPropertyFieldSite[];
  context: WebPartContext
}
