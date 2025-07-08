// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

let config = require(__dirname + "/../config/config.json");

function getAuthHeader(accessToken) {

    // Function to append Bearer against the Access Token
    return "Bearer ".concat(accessToken);
}

function validateConfig() {
    const tenant_id = process.env.tenantId;
    if (!tenant_id) {
        console.log("Tenant id not found");
    }
    const client_secret = process.env.clientSecret;
    if (!client_secret) {
        console.log("client secret not found");
    }
    const client_id = process.env.CLIENTID;
    if (!client_id) {
        console.log("client id not found");
    }    
    const service_principal_id = process.env.servicePrincipalObjectId;
    if (!service_principal_id) {
        console.log("service principal id not found");
    }
    const report_id = process.env.reportId;
    if (!report_id) {
        console.log("report id not found");
    }    
    const workspace_id = process.env.workspaceId;
    if (!workspace_id) {
        console.log("workspace id not found");
    }

    // Validation function to check whether the Configurations are available in the config.json file or not

    let guid = require("guid");

    if (!config.authenticationMode) {
        return "AuthenticationMode is empty. Please choose MasterUser or ServicePrincipal in config.json.";
    }

    if (config.authenticationMode.toLowerCase() !== "masteruser" && config.authenticationMode.toLowerCase() !== "serviceprincipal") {
        return "AuthenticationMode is wrong. Please choose MasterUser or ServicePrincipal in config.json";
    }

    if (!client_id) {
        return "ClientId is empty. Please register your application as Native app in https://dev.powerbi.com/apps and fill Client Id in config.json.";
    }

    if (!guid.isGuid(client_id)) {
        return "ClientId must be a Guid object. Please register your application as Native app in https://dev.powerbi.com/apps and fill Client Id in config.json.";
    }

    if (!report_id) {
        return "ReportId is empty. Please select a report you own and fill its Id in config.json.";
    }

    if (!guid.isGuid(report_id)) {
        return "ReportId must be a Guid object. Please select a report you own and fill its Id in config.json.";
    }

    if (!workspace_id) {
        return "WorkspaceId is empty. Please select a group you own and fill its Id in config.json.";
    }

    if (!guid.isGuid(workspace_id)) {
        return "WorkspaceId must be a Guid object. Please select a workspace you own and fill its Id in config.json.";
    }

    if (!config.authorityUrl) {
        return "AuthorityUrl is empty. Please fill valid AuthorityUrl in config.json.";
    }

    if (config.authenticationMode.toLowerCase() === "masteruser") {
        if (!config.pbiUsername || !config.pbiUsername.trim()) {
            return "PbiUsername is empty. Please fill Power BI username in config.json.";
        }

        if (!config.pbiPassword || !config.pbiPassword.trim()) {
            return "PbiPassword is empty. Please fill password of Power BI username in config.json.";
        }
    } else if (config.authenticationMode.toLowerCase() === "serviceprincipal") {
        if (!client_secret || !client_secret.trim()) {
            return "ClientSecret is empty. Please fill Power BI ServicePrincipal ClientSecret in config.json.";
        }

        if (!tenant_id) {
            return "TenantId is empty. Please fill the TenantId in config.json.";
        }

        if (!guid.isGuid(tenant_id)) {
            return "TenantId must be a Guid object. Please select a workspace you own and fill its Id in config.json.";
        }
    }
}

module.exports = {
    getAuthHeader: getAuthHeader,
    validateConfig: validateConfig,
}