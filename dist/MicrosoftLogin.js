"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
var React = require("react");
var msal_1 = require("msal");
var MicrosoftLoginButton_1 = require("./MicrosoftLoginButton");
var CLIENT_ID_REGEX = /[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}/;
var MicrosoftLogin = (function (_super) {
    __extends(MicrosoftLogin, _super);
    function MicrosoftLogin(props) {
        var _this = _super.call(this, props) || this;
        var debug = props.debug, graphScopes = props.graphScopes, withUserData = props.withUserData;
        var scope = (graphScopes || []);
        scope.some(function (el) { return el.toLowerCase() === "user.read"; }) ||
            scope.push("user.read");
        _this.state = {
            msalInstance: props.clientId &&
                CLIENT_ID_REGEX.test(props.clientId) &&
                new msal_1.UserAgentApplication(props.clientId, null, function () { }, {
                    cacheLocation: "localStorage"
                }),
            scope: scope,
            debug: debug || false,
            withUserData: withUserData || false
        };
        return _this;
    }
    MicrosoftLogin.prototype.componentDidMount = function () {
        var msalInstance = this.state.msalInstance;
        if (msalInstance) {
            this.initialize(msalInstance);
        }
        else {
            this.log("Initialization", "clientID broken or not provided", true);
        }
    };
    MicrosoftLogin.prototype.componentDidUpdate = function (prevProps, prevState) {
        var clientId = this.props.clientId;
        if (prevProps.clientId !== clientId) {
            this.setState({
                msalInstance: clientId &&
                    CLIENT_ID_REGEX.test(clientId) &&
                    new msal_1.UserAgentApplication(clientId, null, function () { })
            });
        }
    };
    MicrosoftLogin.prototype.initialize = function (msalInstance) {
        var _a = this.state, scope = _a.scope, debug = _a.debug, withUserData = _a.withUserData;
        var authCallback = this.props.authCallback;
        if (msalInstance.getUser() &&
            !msalInstance.isCallback(window.location.hash) &&
            window.localStorage.outlook_login_initiated) {
            window.localStorage.removeItem("outlook_login_initiated");
            debug && this.log("Fetch Azure AD 'token' with redirect SUCCEDEED", true);
            debug &&
                this.log("Fetch Graph API 'access_token' in silent mode STARTED", true);
            this.getGraphAPITokenAndUser(msalInstance, scope, withUserData, authCallback, true, debug);
        }
        else if (!msalInstance.isCallback(window.location.hash) &&
            window.localStorage.outlook_login_initiated) {
            window.localStorage.removeItem("outlook_login_initiated");
            debug &&
                this.log("Fetch Azure AD 'token' with redirect FAILED", "Something went wrong", true);
            authCallback("Something went wrong");
        }
    };
    MicrosoftLogin.prototype.login = function () {
        var _a = this.state, msalInstance = _a.msalInstance, scope = _a.scope, withUserData = _a.withUserData, debug = _a.debug;
        var authCallback = this.props.authCallback;
        if (msalInstance) {
            debug && this.log("Login STARTED", true);
            if (this.checkToIE()) {
                this.redirectLogin(msalInstance, scope, debug);
            }
            else {
                this.popupLogin(msalInstance, scope, withUserData, authCallback, debug);
            }
        }
        else {
            this.log("Login FAILED", "clientID broken or not provided", true);
        }
    };
    MicrosoftLogin.prototype.getGraphAPITokenAndUser = function (msalInstance, scope, withUserData, authCallback, isRedirect, debug) {
        var _this = this;
        return msalInstance
            .acquireTokenSilent(scope)
            .catch(function (error) {
            debug &&
                _this.log("Fetch Graph API 'access_token' in silent mode is FAILED", error, true);
            debug &&
                _this.log("Fetch Graph API 'access_token' with " + (isRedirect ? "redirect" : "popup") + " STARTED", true);
            return isRedirect
                ? msalInstance.acquireTokenRedirect(scope)
                : msalInstance.acquireTokenPopup(scope);
        })
            .then(function (accessToken) {
            debug &&
                _this.log("Fetch Graph API 'access_token' SUCCEDEED", accessToken);
            if (withUserData) {
                _this.getUserData(accessToken);
            }
            else {
                debug && _this.log("Login SUCCEDED", true);
                authCallback(null, { accessToken: accessToken });
            }
        })
            .catch(function (error) {
            _this.log("Login FAILED", error, true);
            authCallback(error);
        });
    };
    MicrosoftLogin.prototype.popupLogin = function (msalInstance, scope, withUserData, authCallback, debug) {
        var _this = this;
        debug && this.log("Fetch Azure AD 'token' with popup STARTED", true);
        msalInstance.loginPopup(scope).then(function (idToken) {
            debug && _this.log("Fetch Azure AD 'token' with popup SUCCEDEED", idToken);
            debug &&
                _this.log("Fetch Graph API 'access_token' in silent mode STARTED", true);
            _this.getGraphAPITokenAndUser(msalInstance, scope, withUserData, authCallback, false, debug);
        });
    };
    MicrosoftLogin.prototype.redirectLogin = function (msalInstance, scope, debug) {
        debug && this.log("Fetch Azure AD 'token' with redirect STARTED", true);
        window.localStorage.setItem("outlook_login_initiated", "true");
        msalInstance.loginRedirect(scope);
    };
    MicrosoftLogin.prototype.getUserData = function (token) {
        var _this = this;
        var _a = this.props, authCallback = _a.authCallback, debug = _a.debug;
        debug && this.log("Fetch Graph API user data STARTED", true);
        var options = {
            method: "GET",
            headers: {
                Authorization: "Bearer " + token
            }
        };
        return fetch("https://graph.microsoft.com/v1.0/me", options)
            .then(function (response) { return response.json(); })
            .then(function (userData) {
            debug && _this.log("Fetch Graph API user data SUCCEDEED", userData);
            debug && _this.log("Login SUCCEDED", true);
            authCallback(null, __assign({}, userData, { accessToken: token }));
        });
    };
    MicrosoftLogin.prototype.checkToIE = function () {
        var ua = window.navigator.userAgent;
        var msie = ua.indexOf("MSIE ");
        var msie11 = ua.indexOf("Trident/");
        var msedge = ua.indexOf("Edge/");
        var isIE = msie > 0 || msie11 > 0;
        var isEdge = msedge > 0;
        return isIE || isEdge;
    };
    MicrosoftLogin.prototype.log = function (name, content, isError) {
        var style = "background-color: " + (isError ? "#990000" : "#009900") + "; color: #ffffff; font-weight: 700; padding: 2px";
        console.groupCollapsed("MSLogin debug");
        console.log("%c" + name, style);
        console.log(content);
        console.groupEnd();
    };
    MicrosoftLogin.prototype.render = function () {
        var _a = this.props, buttonTheme = _a.buttonTheme, className = _a.className, customButton = _a.customButton;
        if (customButton) {
            return React.cloneElement(customButton, {
                onClick: this.login.bind(this),
            });
        }
        return (React.createElement("div", null,
            React.createElement(MicrosoftLoginButton_1.default, { buttonTheme: buttonTheme || "light", buttonClassName: className, onClick: this.login.bind(this) })));
    };
    return MicrosoftLogin;
}(React.Component));
exports.default = MicrosoftLogin;
//# sourceMappingURL=MicrosoftLogin.js.map