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
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from "react";
import { HoverCard } from 'office-ui-fabric-react';
import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';
import styles from "./TreeOrgChart.module.scss";
import "react-sortable-tree/style.css";
import { Persona, PersonaSize, } from "office-ui-fabric-react/lib/Persona";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import SPService from "../../../services/SPServices";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/components/Spinner";
import { Web } from "sp-pnp-js";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IconButton } from "office-ui-fabric-react/lib/Button";
import SortableTree from "react-sortable-tree";
//import OrgChart from"@dabeng/react-orgchart"
import Card from '@material-ui/core/Card';
import CardContent from '@material-ui/core/CardContent';
import CardMedia from '@material-ui/core/CardMedia';
import Typography from '@material-ui/core/Typography';
import CardActionArea from '@material-ui/core/CardActionArea';
import Button from '@material-ui/core/Button';
import CardActions from '@material-ui/core/CardActions';
var TreeOrgChart = /** @class */ (function (_super) {
    __extends(TreeOrgChart, _super);
    function TreeOrgChart(props) {
        var _this = _super.call(this, props) || this;
        _this._getPeoplePickerUserItems = function (items) {
            if (items.length > 0) {
                var userEmail = items[0].secondaryText;
                _this.getUserId(userEmail).then(function (userId) {
                    _this.setState({
                        userEmail: userEmail,
                        userId: userId
                    }, function () { _this.loadOrgchart(userEmail); });
                });
            }
            else {
                _this.setState({
                    userEmail: "",
                    userId: 0
                });
            }
        };
        _this.SPService = new SPService(_this.props.context);
        _this.state = {
            treeData: [],
            isLoading: true,
            userEmail: '',
            userId: 0,
        };
        return _this;
    }
    TreeOrgChart.prototype.handleTreeOnChange = function (treeData) {
        this.setState({ treeData: treeData });
    };
    TreeOrgChart.prototype.getUserId = function (email) {
        var web = new Web(this.props.customUrl);
        return web.ensureUser(email).then(function (result) {
            return result.data.Id;
        });
    };
    TreeOrgChart.prototype.componentDidUpdate = function (prevProps, prevState) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!(this.props.currentUserTeam !== prevProps.currentUserTeam ||
                            this.props.maxLevels !== prevProps.maxLevels)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.loadOrgchart(this.props.context.pageContext.user.loginName)];
                    case 1:
                        _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/];
                }
            });
        });
    };
    TreeOrgChart.prototype.componentDidMount = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.loadOrgchart(this.props.context.pageContext.user.loginName)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    TreeOrgChart.prototype.loadOrgchart = function (newValue) {
        return __awaiter(this, void 0, void 0, function () {
            var currentUser, currentUserProperties, treeManagers, treeManagers;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        this.setState({ treeData: [], isLoading: true });
                        currentUser = "i:0#.f|membership|" + newValue;
                        return [4 /*yield*/, this.SPService.getUserProperties(currentUser)];
                    case 1:
                        currentUserProperties = _a.sent();
                        this.treeData = [];
                        if (!!this.props.currentUserTeam) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.buildOrganizationChart(currentUserProperties)];
                    case 2:
                        treeManagers = _a.sent();
                        if (treeManagers)
                            this.treeData.push(treeManagers);
                        return [3 /*break*/, 5];
                    case 3: return [4 /*yield*/, this.buildMyTeamOrganizationChart(currentUserProperties)];
                    case 4:
                        treeManagers = _a.sent();
                        if (treeManagers)
                            this.treeData.push({
                                title: treeManagers.person,
                                expanded: true,
                                children: treeManagers.treeChildren
                            });
                        _a.label = 5;
                    case 5:
                        this.setState({ treeData: this.treeData, isLoading: false });
                        return [2 /*return*/];
                }
            });
        });
    };
    TreeOrgChart.prototype.buildOrganizationChart = function (currentUserProperties) {
        return __awaiter(this, void 0, void 0, function () {
            var treeManagers;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        treeManagers = null;
                        if (!(currentUserProperties.ExtendedManagers &&
                            currentUserProperties.ExtendedManagers.length > 0)) return [3 /*break*/, 2];
                        return [4 /*yield*/, this.getUsers(currentUserProperties.ExtendedManagers[0])];
                    case 1:
                        treeManagers = _a.sent();
                        _a.label = 2;
                    case 2: return [2 /*return*/, treeManagers];
                }
            });
        });
    };
    TreeOrgChart.prototype.getUsers = function (manager) {
        return __awaiter(this, void 0, void 0, function () {
            var person, spUser, managerProperties, imageInitials, classNames4, onRenderCompactCard4, expandingCardProps4, usersDirectReports;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        spUser = {};
                        return [4 /*yield*/, this.SPService.getUserProperties(manager)];
                    case 1:
                        managerProperties = _a.sent();
                        imageInitials = managerProperties.DisplayName.split(" ");
                        spUser.imageUrl = "/_layouts/15/userphoto.aspx?size=L&username=" + managerProperties.Email;
                        spUser.imageInitials = "" + imageInitials[0]
                            .substring(0, 1)
                            .toUpperCase() + imageInitials[1].substring(0, 1).toUpperCase();
                        spUser.text = managerProperties.DisplayName;
                        spUser.tertiaryText = managerProperties.Email;
                        spUser.secondaryText = managerProperties.Title;
                        classNames4 = mergeStyleSets({
                            compactCard: {
                                color: "black",
                                fontWeight: "bold",
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                height: '100%',
                            },
                            item: {
                                selectors: {
                                    '&:hover': {
                                        textDecoration: 'underline',
                                        cursor: 'pointer',
                                    },
                                },
                            },
                        });
                        onRenderCompactCard4 = function () {
                            return (React.createElement("div", { className: classNames4.compactCard },
                                managerProperties.DisplayName,
                                React.createElement("br", null),
                                managerProperties.Title));
                        };
                        expandingCardProps4 = {
                            onRenderCompactCard: onRenderCompactCard4,
                        };
                        person = (React.createElement(HoverCard, { expandingCardProps: expandingCardProps4, instantOpenOnClick: true },
                            React.createElement(Persona, __assign({}, spUser, { hidePersonaDetails: false, size: PersonaSize.size40 }))));
                        if (!(managerProperties.DirectReports &&
                            managerProperties.DirectReports.length > 0)) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.getChildren(managerProperties.DirectReports)];
                    case 2:
                        usersDirectReports = _a.sent();
                        return [2 /*return*/, { id: 1, title: person, expanded: true, children: usersDirectReports }];
                    case 3: return [2 /*return*/, { id: 0, title: person }];
                }
            });
        });
    };
    TreeOrgChart.prototype.getChildren = function (userDirectReports) {
        return __awaiter(this, void 0, void 0, function () {
            var treeChildren, spUser, _loop_1, this_1, _i, userDirectReports_1, user;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        treeChildren = [];
                        spUser = {};
                        _loop_1 = function (user) {
                            var managerProperties, imageInitials, send_email_report, classNames0, onRenderCompactCard0, onRenderExpandedCard0, expandingCardProps0, person, usersDirectReports;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0: return [4 /*yield*/, this_1.SPService.getUserProperties(user)];
                                    case 1:
                                        managerProperties = _a.sent();
                                        imageInitials = managerProperties.DisplayName.split(" ");
                                        spUser.imageUrl = "/_layouts/15/userphoto.aspx?size=L&username=" + managerProperties.Email;
                                        spUser.imageInitials = "" + imageInitials[0]
                                            .substring(0, 1)
                                            .toUpperCase() + imageInitials[1].substring(0, 1).toUpperCase();
                                        spUser.text = managerProperties.DisplayName;
                                        spUser.tertiaryText = managerProperties.Email;
                                        spUser.secondaryText = managerProperties.Title;
                                        send_email_report = "mailto:" + managerProperties.Email;
                                        classNames0 = mergeStyleSets({
                                            compactCard: {
                                                color: "black",
                                                fontWeight: "bold",
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                height: '100%',
                                            },
                                            expandedCard: {
                                                font: "Times New Roman",
                                                color: "black",
                                                fontWeight: "lighter",
                                                marginTop: "20px",
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                marginLeft: "100px",
                                            },
                                            btn: {
                                                marginLeft: "45px",
                                                width: "130px"
                                            },
                                            lnk: {
                                                marginLeft: "10px",
                                                fontFamily: "bold"
                                            },
                                            root: {
                                                height: 395,
                                                maxWidth: 345,
                                            },
                                            media: {
                                                height: 200,
                                            },
                                            item: {
                                                selectors: {
                                                    '&:hover': {
                                                        textDecoration: 'underline',
                                                        cursor: 'pointer',
                                                    },
                                                },
                                            },
                                        });
                                        onRenderCompactCard0 = function () {
                                            return (React.createElement("div", { className: classNames0.compactCard },
                                                managerProperties.DisplayName,
                                                React.createElement("br", null),
                                                managerProperties.Title));
                                        };
                                        onRenderExpandedCard0 = function () {
                                            return (React.createElement(Card, { className: classNames0.root },
                                                React.createElement(CardActionArea, null,
                                                    React.createElement(CardMedia, { className: classNames0.media, image: "/_layouts/15/userphoto.aspx?size=L&username=" + managerProperties.Email, title: managerProperties.DisplayName }),
                                                    React.createElement(CardContent, null,
                                                        React.createElement(Typography, { gutterBottom: true, variant: "h5", component: "h2" }, managerProperties.DisplayName),
                                                        React.createElement(Typography, { variant: "body2", color: "textSecondary", component: "p" },
                                                            managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length > 0 && managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'UserName'; }) ? managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'UserName'; }).Value ? (React.createElement("span", null,
                                                                managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'UserName'; }).Value,
                                                                React.createElement("br", null))) : null : null,
                                                            managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length > 0 && managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }) ? managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value ? (React.createElement("span", null,
                                                                managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value,
                                                                React.createElement("br", null))) : null : null,
                                                            managerProperties && managerProperties.UserProfileProperties && managerProperties.UserProfileProperties.length > 0 && managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }) ? managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value ? (React.createElement("span", null,
                                                                managerProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value,
                                                                React.createElement("br", null))) : null : null))),
                                                React.createElement(CardActions, null,
                                                    React.createElement(Button, { onClick: function () { return _this.loadOrgchart(managerProperties.Email); }, size: "small", color: "primary" }, "Visit OrgChart"),
                                                    React.createElement(Button, { href: send_email_report, size: "small", color: "primary" }, "Send Email"),
                                                    React.createElement(Button, { href: managerProperties.UserUrl, size: "small", color: "primary" }, "Sharepoint"))));
                                        };
                                        expandingCardProps0 = {
                                            onRenderCompactCard: onRenderCompactCard0,
                                            onRenderExpandedCard: onRenderExpandedCard0,
                                            expandedCardHeight: 395
                                        };
                                        person = (React.createElement(HoverCard, { expandingCardProps: expandingCardProps0, instantOpenOnClick: true },
                                            React.createElement(Persona, __assign({}, spUser, { hidePersonaDetails: false, size: PersonaSize.size40 }))));
                                        return [4 /*yield*/, this_1.getChildren(managerProperties.DirectReports)];
                                    case 2:
                                        usersDirectReports = _a.sent();
                                        usersDirectReports
                                            ? treeChildren.push({ title: person, children: usersDirectReports })
                                            : treeChildren.push({ title: person });
                                        return [2 /*return*/];
                                }
                            });
                        };
                        this_1 = this;
                        _i = 0, userDirectReports_1 = userDirectReports;
                        _a.label = 1;
                    case 1:
                        if (!(_i < userDirectReports_1.length)) return [3 /*break*/, 4];
                        user = userDirectReports_1[_i];
                        return [5 /*yield**/, _loop_1(user)];
                    case 2:
                        _a.sent();
                        _a.label = 3;
                    case 3:
                        _i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/, treeChildren];
                }
            });
        });
    };
    TreeOrgChart.prototype.buildMyTeamOrganizationChart = function (currentUserProperties) {
        var _a;
        return __awaiter(this, void 0, void 0, function () {
            var manager, me, treeChildren, imageInitials, hasManager, managerCard, myManager, managerProperties_1, send_email_manager_1, classNames_1, onRenderCompactCard, onRenderExpandedCard, expandingCardProps, meImageInitials, send_email_report, classNames2, onRenderExpandedCard2, onRenderCompactCard2, expandingCardProps2, meCard, usersDirectReports;
            var _this = this;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        manager = {};
                        me = {};
                        treeChildren = [];
                        hasManager = false;
                        return [4 /*yield*/, this.SPService.getUserProfileProperty(currentUserProperties.AccountName, "Manager")];
                    case 1:
                        myManager = _b.sent();
                        if (!myManager) return [3 /*break*/, 3];
                        return [4 /*yield*/, this.SPService.getUserProperties(myManager)];
                    case 2:
                        managerProperties_1 = _b.sent();
                        imageInitials = (_a = managerProperties_1.DisplayName) === null || _a === void 0 ? void 0 : _a.split(" ").map(function (name) { return name[0]; });
                        manager.imageUrl = "/_layouts/15/userphoto.aspx?size=L&username=" + managerProperties_1.Email;
                        if (imageInitials)
                            manager.imageInitials = ("" + imageInitials[0] + imageInitials[1]).toUpperCase();
                        manager.text = managerProperties_1.DisplayName;
                        manager.tertiaryText = managerProperties_1.Email;
                        manager.secondaryText = managerProperties_1.Title;
                        send_email_manager_1 = "mailto:" + managerProperties_1.Email;
                        classNames_1 = mergeStyleSets({
                            compactCard: {
                                color: "black",
                                fontWeight: "bold",
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                height: '100%'
                            },
                            expandedCard: {
                                font: "Times New Roman",
                                color: "black",
                                fontWeight: "lighter",
                                marginTop: "20px",
                                alignItems: 'center',
                                justifyContent: 'center',
                                marginLeft: "100px",
                            },
                            root: {
                                height: 395,
                                maxWidth: 345,
                            },
                            media: {
                                height: 200,
                            },
                            item: {
                                selectors: {
                                    '&:hover': {
                                        textDecoration: 'underline',
                                        cursor: 'pointer',
                                    },
                                },
                            },
                        });
                        onRenderCompactCard = function () {
                            return (React.createElement("div", { className: classNames_1.compactCard },
                                managerProperties_1.DisplayName,
                                React.createElement("br", null),
                                managerProperties_1.Title));
                        };
                        onRenderExpandedCard = function () {
                            return (React.createElement(Card, { className: classNames_1.root },
                                React.createElement(CardActionArea, null,
                                    React.createElement(CardMedia, { className: classNames_1.media, image: manager.imageUrl, title: managerProperties_1.DisplayName }),
                                    React.createElement(CardContent, null,
                                        React.createElement(Typography, { gutterBottom: true, variant: "h5", component: "h2" }, managerProperties_1.DisplayName),
                                        React.createElement(Typography, { variant: "body2", color: "textSecondary", component: "p" },
                                            managerProperties_1.Email,
                                            React.createElement("br", null),
                                            managerProperties_1 && managerProperties_1.UserProfileProperties && managerProperties_1.UserProfileProperties.length > 0 && managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }) ? managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value ? (React.createElement("span", null,
                                                managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value,
                                                React.createElement("br", null))) : null : null,
                                            managerProperties_1 && managerProperties_1.UserProfileProperties && managerProperties_1.UserProfileProperties.length > 0 && managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }) ? managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value ? (React.createElement("span", null,
                                                managerProperties_1.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value,
                                                React.createElement("br", null))) : null : null))),
                                React.createElement(CardActions, null,
                                    React.createElement(Button, { onClick: function () { return _this.loadOrgchart(managerProperties_1.Email); }, size: "small", color: "primary" }, "Visit OrgChart"),
                                    React.createElement(Button, { href: send_email_manager_1, size: "small", color: "primary" }, "Send Email"),
                                    React.createElement(Button, { href: managerProperties_1.UserUrl, size: "small", color: "primary" }, "Visit Sharepoint"))));
                        };
                        expandingCardProps = {
                            onRenderCompactCard: onRenderCompactCard,
                            onRenderExpandedCard: onRenderExpandedCard,
                            expandedCardHeight: 395
                        };
                        managerCard = (React.createElement(HoverCard, { expandingCardProps: expandingCardProps, instantOpenOnClick: true },
                            React.createElement(Persona, __assign({}, manager, { size: PersonaSize.size48, coinSize: 60, hidePersonaDetails: false }))));
                        hasManager = true;
                        _b.label = 3;
                    case 3:
                        meImageInitials = currentUserProperties.DisplayName.split(" ");
                        me.imageUrl = "/_layouts/15/userphoto.aspx?size=L&username=" + currentUserProperties.Email;
                        me.imageInitials = "" + meImageInitials[0]
                            .substring(0, 1)
                            .toUpperCase() + meImageInitials[1].substring(0, 1).toUpperCase();
                        me.text = currentUserProperties.DisplayName;
                        me.tertiaryText = currentUserProperties.Email;
                        me.secondaryText = currentUserProperties.Title;
                        send_email_report = "mailto:" + currentUserProperties.Email;
                        classNames2 = mergeStyleSets({
                            compactCard: {
                                textShadow: "100",
                                color: "black",
                                fontWeight: "bold",
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                height: '100%',
                            },
                            person: {},
                            expandedCard: {
                                marginTop: 20,
                                font: "Times New Roman",
                                color: "black",
                                fontWeight: "lighter",
                                alignItems: 'center',
                                justifyContent: 'center',
                                marginLeft: "50px",
                            },
                            btn: {
                                padding: "20px",
                                marginLeft: "45px",
                            },
                            lnk: {
                                padding: "20px",
                                marginLeft: "10px",
                                fontFamily: "bold"
                            },
                            root: {
                                height: 395,
                                maxWidth: 345,
                            },
                            media: {
                                height: 200,
                            },
                            item: {
                                selectors: {
                                    '&:hover': {
                                        textDecoration: 'underline',
                                        cursor: 'pointer',
                                        alignContent: "center"
                                    },
                                },
                            },
                        });
                        onRenderExpandedCard2 = function () {
                            return (React.createElement(Card, { className: classNames2.root },
                                React.createElement(CardActionArea, null,
                                    React.createElement(CardMedia, { className: classNames2.media, image: "/_layouts/15/userphoto.aspx?size=L&username=" + currentUserProperties.Email, title: currentUserProperties.DisplayName }),
                                    React.createElement(CardContent, null,
                                        React.createElement(Typography, { gutterBottom: true, variant: "h5", component: "h2" }, currentUserProperties.DisplayName),
                                        React.createElement(Typography, { variant: "body2", color: "textSecondary", component: "p" },
                                            currentUserProperties.Email,
                                            React.createElement("br", null),
                                            currentUserProperties && currentUserProperties.UserProfileProperties && currentUserProperties.UserProfileProperties.length > 0 && currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }) ? currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value ? (React.createElement("span", null,
                                                currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'WorkPhone'; }).Value,
                                                React.createElement("br", null))) : null : null,
                                            currentUserProperties && currentUserProperties.UserProfileProperties && currentUserProperties.UserProfileProperties.length > 0 && currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }) ? currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value ? (React.createElement("span", null,
                                                currentUserProperties.UserProfileProperties.find(function (x) { return x.Key == 'Office'; }).Value,
                                                React.createElement("br", null))) : null : null))),
                                React.createElement(CardActions, null,
                                    React.createElement(Button, { onClick: function () { return _this.loadOrgchart(currentUserProperties.Email); }, size: "small", color: "primary" }, "Visit OrgChart"),
                                    React.createElement(Button, { href: send_email_report, size: "small", color: "primary" }, "Send Email"),
                                    React.createElement(Button, { href: currentUserProperties.UserUrl, size: "small", color: "primary" }, "Visit Sharepoint"))));
                        };
                        onRenderCompactCard2 = function () {
                            return (React.createElement("div", { className: classNames2.compactCard },
                                me.text,
                                React.createElement("br", null),
                                me.secondaryText,
                                React.createElement("br", null)));
                        };
                        expandingCardProps2 = {
                            onRenderCompactCard: onRenderCompactCard2,
                            onRenderExpandedCard: onRenderExpandedCard2,
                            expandedCardHeight: 395
                        };
                        meCard = (React.createElement("div", null,
                            React.createElement(HoverCard, { expandingCardProps: expandingCardProps2, instantOpenOnClick: true },
                                React.createElement(Persona, __assign({}, me, { initialsColor: "blue", className: classNames2.person, hidePersonaDetails: false, size: PersonaSize.size48, coinSize: 60 })))));
                        return [4 /*yield*/, this.getChildren(currentUserProperties.DirectReports)];
                    case 4:
                        usersDirectReports = _b.sent();
                        if (hasManager) {
                            treeChildren.push({
                                title: meCard,
                                expanded: true,
                                children: usersDirectReports
                            });
                        }
                        else {
                            treeChildren = usersDirectReports;
                            managerCard = meCard;
                        }
                        return [2 /*return*/, { person: managerCard, treeChildren: treeChildren }];
                }
            });
        });
    };
    TreeOrgChart.prototype.render = function () {
        return (React.createElement("div", { className: styles.treeOrgChart },
            React.createElement(PeoplePicker, { context: this.props.context, titleText: "", personSelectionLimit: 1, showtooltip: true, defaultSelectedUsers: this.state.userEmail ? [this.state.userEmail] : [], selectedItems: this._getPeoplePickerUserItems.bind(this), showHiddenInUI: false, principalTypes: [PrincipalType.User], resolveDelay: 1000 }),
            React.createElement(WebPartTitle, { displayMode: this.props.displayMode, title: this.props.title, updateProperty: this.props.updateProperty }),
            this.state.isLoading ? (React.createElement(Spinner, { size: SpinnerSize.large, label: "Loading Organization Chart ..." })) : null,
            React.createElement("div", { className: styles.treeContainer },
                React.createElement(SortableTree, { treeData: this.state.treeData, onChange: this.handleTreeOnChange.bind(this), canDrag: false, canDrop: false, rowHeight: 120, scaffoldBlockPxWidth: 100, rowDirection: "ltr", orientation: "horizontal", maxDepth: this.props.maxLevels, generateNodeProps: function (rowInfo) { return ({
                        buttons: [
                            React.createElement(IconButton, { disabled: false, checked: false, size: 60, iconProps: { iconName: "ContactInfo" }, title: "Contact Info", ariaLabel: "Contact", onClick: function () {
                                    window.open("https://nam.delve.office.com/?p=" + rowInfo.node.title.props.children.props.tertiaryText + "&v=work");
                                } })
                        ]
                    }); } }))));
    };
    return TreeOrgChart;
}(React.Component));
export default TreeOrgChart;
//# sourceMappingURL=TreeOrgChart.js.map