var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import * as strings from 'TreeOrgChartWebPartStrings';
import TreeOrgChart from './components/TreeOrgChart';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { setup as pnpSetup } from '@pnp/common';
var TreeOrgChartWebPart = /** @class */ (function (_super) {
    __extends(TreeOrgChartWebPart, _super);
    function TreeOrgChartWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    TreeOrgChartWebPart.prototype.onInit = function () {
        pnpSetup({
            spfxContext: this.context
        });
        return Promise.resolve();
    };
    TreeOrgChartWebPart.prototype.render = function () {
        var _this = this;
        var element = React.createElement(TreeOrgChart, {
            title: this.properties.title,
            displayMode: this.displayMode,
            updateProperty: function (value) {
                _this.properties.title = value;
            },
            currentUserTeam: this.properties.currentUserTeam,
            maxLevels: this.properties.maxLevels,
            context: this.context,
            customUrl: this.properties.customUrl
        });
        ReactDom.render(element, this.domElement);
    };
    TreeOrgChartWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(TreeOrgChartWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    TreeOrgChartWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('title', {
                                    label: strings.TitleFieldLabel
                                }),
                                PropertyPaneToggle('currentUserTeam', {
                                    label: strings.CurrentUserTeamFieldLabel
                                }),
                                PropertyFieldNumber("maxLevels", {
                                    key: "numberValue",
                                    label: strings.MaxLevels,
                                    description: strings.MaxLevels,
                                    value: this.properties.maxLevels,
                                    maxValue: 10,
                                    minValue: 1,
                                    disabled: false,
                                }), PropertyPaneTextField('customUrl', {
                                    label: 'Custom Url'
                                }),
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return TreeOrgChartWebPart;
}(BaseClientSideWebPart));
export default TreeOrgChartWebPart;
//# sourceMappingURL=TreeOrgChartWebPart.js.map