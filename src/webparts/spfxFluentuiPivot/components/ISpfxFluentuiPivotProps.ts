import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPropertyFieldSite } from "@pnp/spfx-property-controls/lib/PropertyFieldSitePicker";

export interface ISpfxFluentuiPivotProps {
  description: string;
  site: IPropertyFieldSite[];
  context: WebPartContext
}
