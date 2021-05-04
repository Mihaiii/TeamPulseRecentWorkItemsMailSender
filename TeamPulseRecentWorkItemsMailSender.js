const request = require('request');
const buildQuery = require('odata-query').default;
const nodemailer = require('nodemailer');
const LocalStorage = require('node-localstorage').LocalStorage;
const ls = new LocalStorage('./');

function readConfig() {
    globalObj.config = {
        general: {
            teamPulseUrl: '', //
            teamPulseCredentials: {
                username: '', //
                pass: '' //
            },
            mailServer: {
                required: {
                    server: '', //
                    port: '', //
                    username: '', //
                    pass: '', //
                    fromAddress: '', //
                    nodemailerTlsRejectUnauthorizedOpt : false,
                    nodemailerSecureOpt: false
                },
                toAddress: '' //
            },
            filteredValues: [], //
            filteredProp: "name",
            //depending on what you want the script to do, this could also be () => globalObj.config.general.filteredValues.length
            getNoSpamModeValue: () => globalObj.config.general.mailServer.toAddress,
            adminEmail: '', //
        },
        method: {
            GET: 'GET',
            POST: 'POST'
        },
        mailText: {
            tooManyWorkItems: (exceededNr, remoteSearchID) => 'Ar fi urmat sa se trimita peste {0} de item-uri, deci cel mai probabil ceva nu este ok. ID-ul dupa care s-a facut cautarea este {1}, iar ID-ul primului item adus in urma requestului este {2}'.format(exceededNr, globalObj.lastId, remoteSearchID),
            defaultReminder: 'Mai tii minte aplicatia aia pe care ai facut-o prin 2021 ca sa-ti trimita notificari teampulse?',
            ohShitItIsStillWorking: 'Opreste si tu rularea task-ului ca inca mai merge de nebun.',
            err: (errMsg) => 'A crapat. Oricum probabil n-o mai folosesti, deci poti opri rularea taskului.\nMesajul de eroare este:\n{0}'.format(errMsg),
            errMailSubject: 'TeamPulse script error',
            defaultMailSubject: 'Chestii noi pe TeamPulse alocate tie',
        },
        others: {
            authURLPath: '/Authenticate/WRAPv0.9',
            wrapClientID: 'uri:TeamPulse',
            maxNewWorkitemsWithoutSendingErr: 150,
            maxNrOfDaysInARowWithNoItem: 40,
            mailBodyPreamble: '',
            onlyWithoutWorkCompleted: true
        }
    };
    Object.freeze(globalObj.config);
};

var globalObj = {
    accessToken: '',
    refreshToken: '',
    lastId: '',
    usersData: [{
        id: '',
        name: '',
        firstName: '',
        lastName: '',
        displayName: '',
        email: ''
    }],
    latestWorkItemsData: [{
        id: 0,
        type: "Task",
        projectId: 0,
        fields: {
            AreaID: 0,
            AssignedToID: 0,
            TeamID: 0,
            BacklogPriority: 0,
            Description: '',
            EstimateOptimistic: null,
            EstimatePessimistic: null,
            EstimateProbable: null,
            PertEstimate: null,
            IterationID: 0,
            Name: '',
            ParentID: 0,
            Status: '',
            WorkRemaining: null,
            WorkCompleted: null,
            TfsID: 0,
            tp_DueDate_cf: '',
            tp_Raspuns_cf: null,
            tp_Zi_cf: null,
            tp_Guid_cf: null,
            tp_Tip_cf: null,
        },
        createdBy: '',
        createdAt: '',
        lastModifiedBy: '',
        lastModifiedAt: '',
    }],
    mailTransporterOptions: {},
    mailTransporter: {},
    exit: process.exit
};

var Main = (async function() {

    declareFormatFn();
    readConfig();

    enforceNoSpamMode(globalObj.exit);
    
    await authenticate((q) => setTokens(q)).catch(() => {});
    await createUsersFilter((r) => getUsers(r, (t) => setUsers(t))).catch(() => {});
	await checkAndSetLastTimeWorkItemIdFromLocal(() => getLastTimeWorkItemIdFromRemote((w) => setAndSaveLastTimeWorkItemIdFromRemoteResponse(w))).catch(() => {});
    await createLatestItemsFilter((u) => getLatestItems(u, (j) => setLatestItems(j))).catch(() => {});
    await checkMaxItemsShouldNotExceed((errMsg) => sendErrMailToAdminAndExit(errMsg)).catch(() => {});
    sendMailsWithNewWorkItems(() => saveNewestWorkItemIdAsLastTime());
    checkApproximateLatestWorkItemDate((errMsg) => sendErrMailToAdminAndExit(errMsg), () => saveApproximateLatestWorkItemDate());
    
})();

function makeRequest(urlPath, method, body, processResponseFn, query) {
    if (!urlPath || !method) {
        throw new InvalidArgumentException([urlPath, method]);
    }

    var reqOptions = {
        url: globalObj.config.general.teamPulseUrl + urlPath + (query || ''),
        method: method,
        formData: body
    };

    if (globalObj.accessToken) {
        reqOptions.headers = {
            'Authorization': 'WRAP access_token=' + globalObj.accessToken
        };
    };

    return new Promise((resolve, reject) => {
        request(reqOptions, async function(error, response, body) {
            if(!response || response.statusCode != 200 || error) {
                await sendErrMailToAdminAndExit("\nResponse error or unexpected status code:\nError: {0}\nStautsCode: {1}\nUrl: {2}\nMethod: {3}\nBody: {4}".format(error, response && response.statusCode, reqOptions.url, reqOptions.method, JSON.stringify(reqOptions.formData)));
                reject(); 
            }
            
            var result = "";
            try {
                result = JSON.parse(body);
            } catch (e) {
                result = body;
            }
            
            if (typeof processResponseFn == 'function') {
                processResponseFn(result);
                resolve();
            }
            
            resolve(result);
        });
    });
};

function authenticate(callback) {
    var urlPath = globalObj.config.others.authURLPath;
    var method = globalObj.config.method.POST;
    var body = {
        wrap_client_id: globalObj.config.others.wrapClientID,
        wrap_username: globalObj.config.general.teamPulseCredentials.username,
        wrap_password: globalObj.config.general.teamPulseCredentials.pass
    };

    return makeRequest(urlPath, method, body, callback);
};

function createUsersFilter(callback) {
    var filteredValues = globalObj.config.general.filteredValues || [];
    var filteredProp = globalObj.config.general.filteredProp || 'name';
    var filter = { or: filteredValues.map((elem) => elem = { [filteredProp]: elem })};
    var query = buildQuery({ filter });
    if (typeof callback == 'function') {
        return callback(query);
    };
    return query;
};

function getUsers(filter, callback) {
    var urlPath = '/api/users';
    var method = globalObj.config.method.GET;
    var body = null;
    return makeRequest(urlPath, method, body, callback, filter);
};

function setUsers(data) {
    if (!data) {
        throw new InvalidArgumentException(data);
    }
    
    globalObj.usersData = data.results || data;
};

function setTokens(data, callback) {
    var accessToken = data.match(/wrap_access_token=(.*?)&/)[1];
    var refreshToken = data.match(/wrap_refresh_token=(.*?)&/)[1];

    globalObj.accessToken = decodeURIComponent(accessToken);
    //TODO when will be needed
    globalObj.refreshToken = decodeURIComponent(refreshToken);
};

function checkAndSetLastTimeWorkItemIdFromLocal(callback) {
    globalObj.lastId = ls.getItem('lastId');
    if (typeof globalObj.lastId === "undefined" || globalObj.lastId === null) {
        if (typeof callback == 'function') {
            return callback();
        };
    };
    return Promise.resolve();
};

function getLastTimeWorkItemIdFromRemote(callback) {
    var urlPath = '/api/v1/workitems';
    var method = globalObj.config.method.GET;
    var body = null;
    var top = 1;
    var orderBy = ['id desc']
    return makeRequest(urlPath, method, body, callback, buildQuery({ orderBy, top }));
};

function setAndSaveLastTimeWorkItemIdFromRemoteResponse(data) {
    if (!data) {
        throw new InvalidArgumentException(data);
    }
    
    var info = data.results || data;
    globalObj.lastId = info[0].id;
    ls.setItem('lastId', globalObj.lastId);
};

function saveNewestWorkItemIdAsLastTime() {
    if (globalObj.latestWorkItemsData.length && globalObj.latestWorkItemsData[0].id) {
        ls.setItem('lastId', globalObj.latestWorkItemsData[0].id);
    }
};

function enforceNoSpamMode(callback) {
    if (typeof callback != 'function') {
        throw new InvalidArgumentException(callback);
    }

    if (!globalObj.config.general.getNoSpamModeValue()) {
        callback();
    };
};

function createLatestItemsFilter(callback) {
    var filter = {};
    if (parseInt(globalObj.lastId)) {
        filter.id = { gt: parseInt(globalObj.lastId) };
    }
    
    if(globalObj.usersData.length  && globalObj.config.general.filteredValues.length) { 
        var filteredValues = globalObj.usersData.map(e => e.id);
        var filteredProp = 'AssignedToID';
        filter.or = filteredValues.map((elem) => elem = { [filteredProp]: elem });
    };
    
    if(globalObj.config.others.onlyWithoutWorkCompleted) {
        filter.WorkCompleted = null;
    }
    
    var query = buildQuery({ filter });
    if (typeof callback == 'function') {
        return callback(query);
    };
    return query;
};

function getLatestItems(filter, callback) {
    var urlPath = '/api/v1/workitems';
    var method = globalObj.config.method.GET;
    var body = null;
    return makeRequest(urlPath, method, body, callback, filter);
};

function checkMailTransporterOptions(callback) {
    var result = globalObj.mailTransporterOptions && Object.keys(globalObj.mailTransporterOptions).length !== 0 &&
        globalObj.mailTransporter && Object.keys(globalObj.mailTransporter).length !== 0;

    if (!result && typeof callback == 'function') {
        callback();
    };
    return result;
};

function setMailTransporter() {
    globalObj.mailTransporterOptions = {
        host: globalObj.config.general.mailServer.required.server,
        port: globalObj.config.general.mailServer.required.port,
        secure: globalObj.config.general.mailServer.required.nodemailerSecureOpt,
        auth: {
            user: globalObj.config.general.mailServer.required.username,
            pass: globalObj.config.general.mailServer.required.pass
        },
        tls: {
            rejectUnauthorized: globalObj.config.general.mailServer.required.nodemailerTlsRejectUnauthorizedOpt
        }
    };
    globalObj.mailTransporter = nodemailer.createTransport(globalObj.mailTransporterOptions);
}

function checkMaxItemsShouldNotExceed(validationFailedFn, validationPassedFn, data) {
    
    data = data || globalObj.latestWorkItemsData;
    if (typeof validationFailedFn != 'function' || !data) {
        throw new InvalidArgumentException([validationFailedFn, data]);
    }
    
    var info = data.results || data;
    if (info.length) {
        var usersId = globalObj.usersData.map((el) => el.id);
        var nrOfItemsAboutToBeSent = info.filter((e) => usersId.includes(e.fields.AssignedToID)).length;
        var maxItemsAllowed = globalObj.config.others.maxNewWorkitemsWithoutSendingErr;
        if (nrOfItemsAboutToBeSent > maxItemsAllowed) {
            return validationFailedFn(globalObj.config.mailText.tooManyWorkItems(maxItemsAllowed, info[0].id));
        }
    }
    
    if(typeof validationPassedFn == 'function') {
        return validationPassedFn(info);
    }
    
    return Promise.resolve();
}

function checkApproximateLatestWorkItemDate(validationFailedFn, validationPassedFn, data) {
    if (typeof validationFailedFn != 'function') {
        throw new InvalidArgumentException(validationFailedFn);
    }
    
    var info = parseInt(data) || parseInt(ls.getItem('approxLatestWorkItemDate'));

    var maxNrOfDaysInARowWithNoItem = globalObj.config.others.maxNrOfDaysInARowWithNoItem || 40;

    var someDaysFromLastItem = info + 1000 * 60 * 60 * 24 * parseInt(maxNrOfDaysInARowWithNoItem);

    if (info && someDaysFromLastItem < Date.now()) {
        var ohShitItIsStillWorking = globalObj.config.mailText.ohShitItIsStillWorking;
        return validationFailedFn('{0}'.format(ohShitItIsStillWorking));
    }
    
    if (typeof validationPassedFn == 'function') {
        return validationPassedFn(info);
    }
    
}

function saveApproximateLatestWorkItemDate() {
    if(globalObj.latestWorkItemsData.length != 0) {
        ls.setItem('approxLatestWorkItemDate', Date.now());
    }
}

function setLatestItems(data) {
    if (!data) {
        throw new InvalidArgumentException(data);
    }
    
    globalObj.latestWorkItemsData = data.results || data;
}

function sendMailsWithNewWorkItems(callback, data) {
    if (!data) {
        data = globalObj.latestWorkItemsData;
    }

    var info = data.results || data;
    var promises = [];
    if (info.length) {
        for (var user of globalObj.usersData) {
            var mailBody = globalObj.config.others.mailBodyPreamble || '';
            var teamPulseUrl = globalObj.config.general.teamPulseUrl;

            var itemsAssignedToID = info.filter((e) => e.fields.AssignedToID == user.id);
            if (itemsAssignedToID.length > 0) {
                itemsAssignedToID.map((el) => mailBody += '\n{0}/view#item/{1} - {2}'.format(teamPulseUrl, el.id, el.fields.Name));
                mailInfoObj = {
                    msg: mailBody,
                    toAddress: user.email
                };
                promises.push(sendMail(mailInfoObj));
            }
        };
        Promise.all(promises)
        .then(typeof callback == 'function' && callback())
        .catch((msg) => {sendErrMailToAdminAndExit(msg)});
    };
};

function sendErrMailToAdmin(message) {
    var defaultReminder = globalObj.config.mailText.defaultReminder;
    var errMailSubject = globalObj.config.mailText.errMailSubject;
    var info = globalObj.config.mailText.err(new Error(message).stack);
    return sendMail({
        msg: '{0}\n{1}'.format(defaultReminder, info),
        subject: errMailSubject
    }, true);
}

function sendErrMailToAdminAndExit(message) {
    return new Promise(async function(resolve, reject) {
        await sendErrMailToAdmin(message).catch(() => {});
        //.catch((msg) => console.log(msg));
        resolve();
        globalObj.exit();
    }).catch(() => {});
}

function sendMail(mailInfoObj, sendToAdmin) {
    checkMailTransporterOptions(() => setMailTransporter());
    var mailOptions = {
        from: globalObj.config.general.mailServer.required.fromAddress,
        to: sendToAdmin ? globalObj.config.general.adminEmail : (globalObj.config.general.mailServer.toAddress || mailInfoObj.toAddress),
        subject: mailInfoObj.subject || globalObj.config.mailText.defaultMailSubject,
        text: mailInfoObj.msg
    };
    if (mailOptions.from && mailOptions.to) {
        return new Promise(function(resolve, reject) {
            globalObj.mailTransporter.sendMail(mailOptions, function(error, info) {
                if (error) {
                    reject(error);
                } else {
                    resolve();
                }
            });
        });
    };
};

//https://gist.github.com/justmoon/15511f92e5216fa2624b
function InvalidArgumentException(message, extra) {
  Error.captureStackTrace(this, this.constructor);
  this.name = this.constructor.name;
  this.message = message;
  this.extra = extra;
};

//https://stackoverflow.com/a/4673436/8151327
function declareFormatFn() {
    if (!String.prototype.format) {
        String.prototype.format = function() {
            var args = arguments;
            return this.replace(/{(\d+)}/g, function(match, number) {
                return typeof args[number] != 'undefined' ?
                    args[number] :
                    match;
            });
        };
    };
};