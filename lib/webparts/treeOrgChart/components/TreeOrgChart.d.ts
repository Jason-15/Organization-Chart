import * as React from "react";
import { ITreeOrgChartProps } from "./ITreeOrgChartProps";
import { ITreeOrgChartState } from "./ITreeOrgChartState";
import "react-sortable-tree/style.css";
import { ITreeData } from "./ITreeData";
export default class TreeOrgChart extends React.Component<ITreeOrgChartProps, ITreeOrgChartState> {
    private treeData;
    private SPService;
    constructor(props: any);
    private handleTreeOnChange;
    getUserId(email: string): Promise<number>;
    _getPeoplePickerUserItems: (items: any[]) => void;
    componentDidUpdate(prevProps: ITreeOrgChartProps, prevState: ITreeOrgChartState): Promise<void>;
    componentDidMount(): Promise<void>;
    loadOrgchart(newValue: any): Promise<void>;
    buildOrganizationChart(currentUserProperties: any): Promise<ITreeData | null>;
    private getUsers;
    private getChildren;
    private buildMyTeamOrganizationChart;
    render(): React.ReactElement<ITreeOrgChartProps>;
}
//# sourceMappingURL=TreeOrgChart.d.ts.map