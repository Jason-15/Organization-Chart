import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ITreeOrgChartProps {
    title: string;
    currentUserTeam: boolean;
    maxLevels: number;
    displayMode: DisplayMode;
    updateProperty: (value: string) => void;
    context: WebPartContext;
    customUrl: string;
}
//# sourceMappingURL=ITreeOrgChartProps.d.ts.map