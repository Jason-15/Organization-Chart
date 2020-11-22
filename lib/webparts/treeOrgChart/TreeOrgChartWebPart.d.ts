import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from '@microsoft/sp-webpart-base';
export interface ITreeOrgChartWebPartProps {
    title: string;
    currentUserTeam: boolean;
    maxLevels: number;
    customUrl: string;
}
export default class TreeOrgChartWebPart extends BaseClientSideWebPart<ITreeOrgChartWebPartProps> {
    onInit(): Promise<void>;
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=TreeOrgChartWebPart.d.ts.map