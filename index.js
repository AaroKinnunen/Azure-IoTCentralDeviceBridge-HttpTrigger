/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

//decoder function decodes the hex message and transforms the payload JSON-body
function decoder(payload) {
    
    var constants = {
        'commonMessageHeaderSpr':
        {
            'type':
            {
                01: "Application 3P",
                02: "Application 1P",
                03: "Application VMUMC-OC",
                255: "Application Error"
            },

        },
        'Error_message': {
            1: "Autoscan in progress",
            2: "No meter connected",
            3: "Invalid meter",
            4: "Invalid meter setup"
        },
    };

    /**
     * Start decoder
     * @type {any | undefined}
     */
    var payLoadJson = decodeToJson(payload);
    var result = setPayload();
    var payload_hex;
    var payloadByte = [];
    var telemetryC;
    var devEuiLink;


    /** Helper functions **/
    /**
     * Parsing metadata and payload_hex
     * @param payLoadJson
     * @returns {null|{payload: {device:{deviceId:string} , measurements: {}}}}
     */
    function setPayload() {
        if (payLoadJson !== null) {
            if (payLoadJson !== null && payLoadJson.length !== 0) {
                if (payLoadJson.hasOwnProperty('DevEUI_uplink')) {
                    devEuiLink = payLoadJson.DevEUI_uplink;
                } else if (payLoadJson.hasOwnProperty('DevEUI_downlink_Sent')) {
                    devEuiLink = payLoadJson.DevEUI_downlink_Sent; // just an example if downlinks would be used
                }
                if (devEuiLink !== null && devEuiLink.length !== 0) {
                    var payloadResult = getPayload();
                    if (devEuiLink.hasOwnProperty('payload_hex')) {
                        payload_hex = devEuiLink['payload_hex'];
                        if (payload_hex !== null && payload_hex.length !== 0) {
                            payload_hex = payload_hex.trim();
                            payloadByte = hexStringToBytes(payload_hex);
                            if (payLoadJson.hasOwnProperty('DevEUI_uplink')) {
                                payloadResult.measurements = getTelemetryUp();
                            }
                        }
                    }
                    else {
                        //payloadResult.telemetry =getTelemetryDown();
                        //Downlinks not supported in Device bridge so no need for downlink function
                    }
                    // for the sent downLink
                    //telemetryC.values['sentPayloadHex'] = "00";
                    return payloadResult;
                }
            }
        }
        return null;
    }
    /**
     * Receiving telemetry by value message_type
     * @param payload_hex
     * @param payLoadJson
     * @returns {null|Telemetry }
     */
    function getTelemetryUp() {
        var message_type = payloadByte[0];
        if (message_type > 0) {
            //telemetryC = getTelemetryCommonUp();
        }
        switch (message_type) {
            case 01: // "Application 3P"
                return getTelemetryCarlo_3P();
            case 02: //"Application 1P"
                return getpackets();
            case 03: // "Application VMUMC-OC"
                return null;
            case 255: // Application Error
                return getErrorMessage();
            default:
                return null;
        }
    }
    function getpackets() {
        switch (payloadByte[1]) {
            case 97:
                return getTelemetryCarlo_1P61();
            case 99:
                return getTelemetryCarlo_1P63();
            default:
                return null;
        }
    }



    function getTelemetryCarlo_1P63() {
        var tele = get_1P_Telemetry();
        var mes = hexStringToBytes(payload_hex);
        //tele['kWh (-) TOT']=mes;
        if (mes[1] == 99) {
            var x = mes.slice(2, 6);
            for (var j = 0; j < x.length; j++) {
                tele['kW_L1'] += x[j] << (8 * (j));
            }
            tele['kW_L1'] = tele['kW_L1'] * 0.000001;
            mes.splice(1, 5)

        }
        if (mes[1] == 100) {
            var x1 = mes.slice(2, 6);
            for (var j = 0; j < x1.length; j++) {
                tele['kVA_L1'] += x1[j] << (8 * (j));
            }
            tele['kVA_L1'] = tele['kVA_L1'] * 0.000001;
            mes.splice(1, 5)

        }
        if (mes[1] == 101) {
            var x2 = mes.slice(2, 6);
            for (var j = 0; j < x2.length; j++) {
                tele['kvar_L1'] += x2[j] << (8 * (j));
            }
            tele['kvar_L1'] = tele['kvar_L1'] * 0.000001;
            mes.splice(1, 5)

        }
        if (mes[1] == 60) {
            var x2 = mes.slice(2, 10);
            for (var j = 0; j < x2.length; j++) {
                tele['kWhTOT'] += x2[j] << (8 * (j));
            }
            tele['kWhTOT'] = tele['kWhTOT'] * 0.1;
            mes.splice(1, 9)
        }
        if (mes[1] == 62) {
            var x2 = mes.slice(2, 10);
            for (var j = 0; j < x2.length; j++) {
                tele['kvarhTOT'] += x2[j] << (8 * (j));
            }
            tele['kvarhTOT'] = tele['kvarhTOT'] * 0.1;
            mes.splice(1, 9)
        }
        if (mes[1] == 255) {
            var x = mes.slice(2, 8);
            var d = new Date(...x);
            tele['TimestampEM24_UTC'] = d;
        }



        clean(tele);
        return tele;
    }
    function getTelemetryCarlo_1P61() {
        var tele = get_1P_Telemetry();
        var mes = hexStringToBytes(payload_hex);
        if (mes[1] == 97) {
            var x = mes.slice(2, 6);
            for (var j = 0; j < x.length; j++) {
                tele['V_L1N'] += x[j] << (8 * (j));
            }
            tele['V_L1N'] = tele['V_L1N'] * 0.1;
            mes.splice(1, 5)

        }
        if (mes[1] == 98) {
            var x1 = mes.slice(2, 6);
            for (var j = 0; j < x1.length; j++) {
                tele['A_L1'] += x1[j] << (8 * (j));
            }
            tele['A_L1'] = tele['A_L1'] * 0.001;
            mes.splice(1, 5)

        }
        if (mes[1] == 102) {
            var x2 = mes.slice(2, 4);
            for (var j = 0; j < x2.length; j++) {
                tele['PF_L1'] += x2[j] << (8 * (j));
            }
            tele['PF_L1'] = tele['PF_L1'] * 0.01;
            mes.splice(1, 3)

        }
        if (mes[1] == 103) {
            var x2 = mes.slice(2, 4);
            for (var j = 0; j < x2.length; j++) {
                tele['Hz'] += x2[j] << (8 * (j));
            }
            tele['Hz'] = tele['Hz'] * 0.1;
            mes.splice(1, 3)
        }
        if (mes[1] == 255) {
            var x = mes.slice(2, 8);
            var d = new Date(...x);
            tele['TimestampEM24_UTC'] = d;
        }



        clean(tele);
        return tele;
    }



    function clean(obj) {
        for (var propName in obj) {
            if (obj[propName] === null || obj[propName] === undefined || obj[propName].length === 0) {
                delete obj[propName];
            }

        }
        return obj

    }
    function getTelemetryCarlo_3P() {
        var telemetry3p = get_3P_telemetry();
        var mes = payload_hex.slice(2, payload_hex.length);
        var splitted = mes.match(/.{1,18}/g);
        for (var i = 0; i < splitted.length; i++) {
            var x = hexStringToBytes(splitted[i]);
            switch (x[0]) {
                case 60:
                    var kwhBytes = x.slice(1, 10);

                    for (var j = 0; j < kwhBytes.length; j++) {
                        telemetry3p['kWh (+) TOT (3P)'] += kwhBytes[j] << (8 * (j));
                    }
                    break;
                case 63:
                    var m = x.slice(1, 10);
                    for (var j = 0; j < m.length; j++) {
                        telemetry3p['kWh (-) TOT (3P)'] += m[j] << (8 * (j));
                    }

                    break;
                case 66:
                    var b = x.slice(1, 10);
                    for (var j = 0; j < b.length; j++) {
                        telemetry3p['kWh (+) PAR (3P)'] += b[j] << (8 * (j));
                    }
                    break;
                case 69:
                    var m1 = x.slice(1, 10);
                    for (var j = 0; j < m1.length; j++) {
                        telemetry3p['kWh (-) PAR (3P)'] += m1[j] << (8 * (j));
                    }
                    break;
                case 255:
                    var timestampBytes = x.slice(1, 8);
                    //timestampBytes[0]=123;
                    var d = new Date(...timestampBytes);
                    telemetry3p["TimestampEM24(UTC) (3P)"] = d;
                    break;
                default:
                    break;

            }
        }
        clean(telemetry3p);
        return telemetry3p;
    }
    function get_1P_Telemetry() {
        return {
            'kWhTOT': null,
            'kVAhTOT': null,
            'kvarhTOT': null,

            'kWh (-) TOT': null,
            'kVAh (-) TOT': null,
            'kvarh (-) TOT': null,

            'kWh (+) PAR': null,
            'kVAh (+) PAR': null,
            'kvarh (+) PAR': null,

            'kWh (-) PAR': null,
            'kVAh (-) PAR': null,
            'kvarh (-) PAR': null,

            'kWh (+) t1': null,
            'kWh (+) t2': null,
            'kWh (+) t3': null,
            'kWh (+) t4': null,
            'kWh (+) t5': null,
            'kWh (+) t6': null,

            'kWh (-) t1': null,
            'kWh (-) t2': null,
            'kWh (-) t3': null,
            'kWh (-) t4': null,
            'kWh (-) t5': null,
            'kWh (-) t6': null,

            'kvarh (+) t1': null,
            'kvarh (+) t2': null,
            'kvarh (+) t3': null,
            'kvarh (+) t4': null,
            'kvarh (+) t5': null,
            'kvarh (+) t6': null,

            'kvarh (-) t1': null,
            'kvarh (-) t2': null,
            'kvarh (-) t3': null,
            'kvarh (-) t4': null,
            'kvarh (-) t5': null,
            'kvarh (-) t6': null,

            'V_L1N': null,
            'A_L1': null,
            'kW_L1': null,
            'kVA_L1': null,
            'kvar_L1': null,
            'PF_L1': null,
            'Hz': null,
            'THD V L1-N': null,
            'THD A L1': null,

            'kWh (+) L1 TOT': null,
            'kVAh (+) L1 TOT': null,
            'kvarh (+) L1 TOT': null,

            'kWh (-) L1 TOT': null,
            'kVAh (-) L1 TOT': null,
            'kvarh (-) L1 TOT': null,

            'kWh (+) L1 PAR': null,
            'kVAh (+) L1 PAR': null,
            'kvarh (+) L1 PAR': null,

            'kWh (-) L1 PAR': null,
            'kVAh (-) L1 PAR': null,
            'kvarh (-) L1 PAR': null,
            'TimestampEM24_UTC': null,

        };
    }
    function get_3P_telemetry() {
        return {
            'kWh (+) TOT (3P)': null,
            'kWh (-) TOT (3P)': null,
            'kWh (+) PAR (3P)': null,
            'kWh (-) PAR (3P)': null,
            'kWh (+) t1 (3P)': null,
            'kWh (+) t2 (3P)': null,
            'kWh (+) t3 (3P)': null,
            'kWh (+) t4 (3P)': null,
            'kWh (+) t5 (3P)': null,
            'kWh (+) t6 (3P)': null,
            'kWh (-) t1 (3P)': null,
            'kWh (-) t2 (3P)': null,
            'kWh (-) t3 (3P)': null,
            'kWh (-) t4 (3P)': null,
            'kWh (-) t5 (3P)': null,
            'kWh (-) t6 (3P)': null,
            'kWh (+) L1 TOT (3P)': null,
            'kWh (-) L1 TOT (3P)': null,
            'kWh (+) L1 PAR (3P)': null,
            'kWh (-) L1 PAR (3P)': null,
            'kWh (+) L2 TOT (3P)': null,
            'kWh (-) L2 TOT (3P)': null,
            'kWh (+) L2 PAR (3P)': null,
            'kWh (-) L2PAR (3P)': null,
            'kWh (+) L3 TOT (3P)': null,
            'kWh (-) L3 TOT (3P)': null,
            'kWh (+) L3 PAR (3P)': null,
            'kWh (-) L3 PAR (3P)': null,
            'TimestampEM24(UTC) (3P)': null

        };
    }

    function getErrorMessage() {
        var error_message = payloadByte[1];
        var err;
        switch (error_message) {
            case 01:
                err = constants.Error_message[1];
                break;
            case 02:
                err = constants.Error_message[2];
                break;
            case 03:
                err = constants.Error_message[3];
                break;
            case 04:
                err = constants.Error_message[4];
                break;
            default:
                err = "Unknown error";
                break;
        }
        return {
            'Error': err
        };
    }




    /**
     * Metadata
     * @returns Headers payload: {device: {deviceId: string}, measurements: {}}
     */
    function getPayload() {
        return {
            "device": {
                "deviceId": payLoadJson.DevEUI_uplink.DevEUI
            },
            "measurements": {
            }

        }
    }


    function hexStringToBytes(str) {
        var array = str.match(/.{1,2}/g);
        var a = [];
        array.forEach(function (element) {
            a.push(parseInt(element, 16));
        });
        return a;
    }

    function convertDateTimeISO8601_toMsUtc(str) {
        return new Date(str).getTime();
    }

    function bytesToInt(bytes) {
        var val = 0;
        for (var j = 0; j < bytes.length; j++) {
            val += bytes[j];
            if (j < bytes.length - 1) {
                val = val << 8;
            }
        }
        return val;
    }

    function decodeToJson(payload) {
        try {
            return JSON.parse(String.fromCharCode.apply(String, payload));
        } catch (e) {
            return JSON.parse(JSON.stringify(payload));
        }
    }

    function getBit(byte, bitNumber) {
        return ((byte & (1 << bitNumber)) != 0) ? 1 : 0;
    }

    return result;
}

const fetch = require('node-fetch');
const handleMessage = require('./lib/engine');

const msiEndpoint = process.env.MSI_ENDPOINT;
const msiSecret = process.env.MSI_SECRET;

const parameters = {
    idScope: process.env.ID_SCOPE,
    primaryKeyUrl: process.env.IOTC_KEY_URL
};

let kvToken;

module.exports = async function (context, req) {
    var bod = decoder(req.body);// call the decoder function
    try {
        req.body = bod; // change the message body to required form
        await handleMessage({ ...parameters, log: context.log, getSecret: getKeyVaultSecret }, req.body.device, req.body.measurements, req.body.timestamp);
    } catch (e) {
        context.log('[ERROR]', e.message);

        context.res = {
            status: e.statusCode ? e.statusCode : 500,
            body: e.message
        };
    }
}

/**
 * Fetches a Key Vault secret. Attempts to refresh the token on authorization errors.
 */
async function getKeyVaultSecret(context, secretUrl, forceTokenRefresh = false) {
    if (!kvToken || forceTokenRefresh) {
        const url = `${msiEndpoint}/?resource=https://vault.azure.net&api-version=2017-09-01`;
        const options = {
            method: 'GET',
            headers: { 'Secret': msiSecret }
        };

        try {
            context.log('[HTTP] Requesting new Key Vault token');
            const response = await fetch(url, options).then(res => res.json())
            kvToken = response.access_token;
        } catch (e) {
            context.log('fail: ' + e);
            throw new Error('Unable to get Key Vault token');
        }
    }

    url = `${secretUrl}?api-version=2016-10-01`;
    var options = {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${kvToken}` },
    };

    try {
        context.log('[HTTP] Requesting Key Vault secret', secretUrl);
        const response = await fetch(url, options).then(res => res.json())
        return response && response.value;
    } catch (e) {
        if (e.statusCode === 401 && !forceTokenRefresh) {
            return await getKeyVaultSecret(context, secretUrl, true);
        } else {
            throw new Error('Unable to fetch secret');
        }
    }
}