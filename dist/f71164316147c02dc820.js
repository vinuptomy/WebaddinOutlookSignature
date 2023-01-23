function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }

function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global OfficeRuntime, require */
var documentHelper = require("./documentHelper");

var fallbackAuthHelper = require("./fallbackAuthHelper");

var sso = require("office-addin-sso");

var retryGetAccessToken = 0;
export function getGraphData() {
  return _getGraphData.apply(this, arguments);
}

function _getGraphData() {
  _getGraphData = _asyncToGenerator( /*#__PURE__*/regeneratorRuntime.mark(function _callee() {
    var bootstrapToken, exchangeResponse, mfaBootstrapToken, userprofileresponse, spdataresponse;
    return regeneratorRuntime.wrap(function _callee$(_context) {
      while (1) {
        switch (_context.prev = _context.next) {
          case 0:
            _context.prev = 0;
            _context.next = 3;
            return OfficeRuntime.auth.getAccessToken({
              allowSignInPrompt: true
            });

          case 3:
            bootstrapToken = _context.sent;
            _context.next = 6;
            return sso.getGraphToken(bootstrapToken);

          case 6:
            exchangeResponse = _context.sent;

            if (!exchangeResponse.claims) {
              _context.next = 12;
              break;
            }

            _context.next = 10;
            return OfficeRuntime.auth.getAccessToken({
              authChallenge: exchangeResponse.claims
            });

          case 10:
            mfaBootstrapToken = _context.sent;
            exchangeResponse = sso.getGraphToken(mfaBootstrapToken);

          case 12:
            if (!exchangeResponse.error) {
              _context.next = 16;
              break;
            }

            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            handleAADErrors(exchangeResponse);
            _context.next = 26;
            break;

          case 16:
            _context.next = 18;
            return sso.makeGraphApiCall(exchangeResponse.access_token);

          case 18:
            userprofileresponse = _context.sent;
            documentHelper.writeDataToOfficeDocument(userprofileresponse);
            _context.next = 22;
            return sso.makeGraphApiCallSP(exchangeResponse.access_token);

          case 22:
            spdataresponse = _context.sent;
            console.log(spdataresponse);
            documentHelper.writeOfficeTimeData(spdataresponse);
            sso.showMessage("Your data has been added to the document.");

          case 26:
            _context.next = 31;
            break;

          case 28:
            _context.prev = 28;
            _context.t0 = _context["catch"](0);

            if (_context.t0.code) {
              if (sso.handleClientSideErrors(_context.t0)) {
                fallbackAuthHelper.dialogFallback();
              }
            } else {
              sso.showMessage("EXCEPTION: " + JSON.stringify(_context.t0));
            }

          case 31:
          case "end":
            return _context.stop();
        }
      }
    }, _callee, null, [[0, 28]]);
  }));
  return _getGraphData.apply(this, arguments);
}

function handleAADErrors(exchangeResponse) {
  // On rare occasions the bootstrap token is unexpired when Office validates it,
  // but expires by the time it is sent to AAD for exchange. AAD will respond
  // with "The provided value for the 'assertion' is not valid. The assertion has expired."
  // Retry the call of getAccessToken (no more than once). This time Office will return a
  // new unexpired bootstrap token.
  if (exchangeResponse.error_description.indexOf("AADSTS500133") !== -1 && retryGetAccessToken <= 0) {
    retryGetAccessToken++;
    getGraphData();
  } else {
    fallbackAuthHelper.dialogFallback();
  }
}