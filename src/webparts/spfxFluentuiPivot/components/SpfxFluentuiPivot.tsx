import * as React from 'react';
import styles from './SpfxFluentuiPivot.module.scss';
import { ISpfxFluentuiPivotProps } from './ISpfxFluentuiPivotProps';
import { ISpfxFluentuiPivotState } from './ISpfxFluentuiPivotState';
import { SPService } from '../../service/SPService';
import * as strings from 'SpfxFluentuiPivotWebPartStrings';
import { PivotItem, IPivotItemProps, Pivot } from 'office-ui-fabric-react/lib/Pivot';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { IStyleSet } from 'office-ui-fabric-react/lib/Styling';

const labelStyles: Partial<IStyleSet<ILabelStyles>> = {
  root: { marginTop: 10, marginLeft: 10 },
};

export default class SpfxFluentuiPivot extends React.Component<ISpfxFluentuiPivotProps, ISpfxFluentuiPivotState> {

  private _spService: SPService;

  constructor(props: ISpfxFluentuiPivotProps) {
    super(props);
    this.state = {
      siteUsres: [],
      siteGroups: [],
    }
    this._spService = new SPService(this.props.context);
  }

  public componentDidMount() {
    this.getSiteGroups();
    this.getSiteUsres();
  }

  public async getSiteUsres() {
    if (this.props.site && this.props.site.length) {
      let siteUrl = this.props.site[0].url;
      let users = await this._spService.getCurrentSiteUsers(siteUrl);
      this.setState({ siteUsres: users });
    }
  }

  public async getSiteGroups() {
    if (this.props.site && this.props.site.length) {
      let siteUrl = this.props.site[0].url;
      let groups = await this._spService.getCurrentSiteGroups(siteUrl);
      this.setState({ siteGroups: groups });
    }
  }

  public componentDidUpdate(prevProps: ISpfxFluentuiPivotProps) {
    if (prevProps.site !== this.props.site) {
      this.getSiteGroups();
      this.getSiteUsres();
    }
  }

  public render(): React.ReactElement<ISpfxFluentuiPivotProps> {
    return (
      <React.Fragment>
        {
          this.props.site && this.props.site.length ?
            <React.Fragment>
              <Label>Selected Site: {this.props.site[0].title}</Label>
              <Pivot aria-label="Count and Icon Pivot Example">
                <PivotItem headerText={strings.SiteUserLabel}>
                  {this.state.siteUsres.map((value) =>
                    <Label styles={labelStyles}>{value.Title}</Label>
                  )}
                </PivotItem>
                <PivotItem headerText={strings.SiteGroupsLabel}>
                  {this.state.siteGroups.map((value) =>
                    <Label styles={labelStyles}>{value.Title}</Label>
                  )}
                </PivotItem>
              </Pivot>
            </React.Fragment>
            : <Label>{strings.ErrorMessageLabel}</Label>}
      </React.Fragment>
    );
  }
}
