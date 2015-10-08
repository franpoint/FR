//! FR.Util - Fran Rodriguez's Utilities for SP Online

/// <reference path="lib/underscore-1.6.0.min.js" />
/// <reference path="lib/underscore.string-2.4.0.min.js" />
/// <reference path="lib/jquery-1.11.1.min.js" />


//Usage:
//   namespace("FR.Util") - Defines an empty namespace
//   namespace("FR.Repository", { f1: function(){}, f2: function(){} }); - Initialise the namespace with the supplied object
//   namespace("FR.Util.String", function() { ... return { f1: function() {}, ... } }); - Initialise the namespace with the object returned by the supplied function
var namespace = namespace || function (ns_string, obj) {
    var parts = ns_string.split('.'),
        parent = window;

    _.each(parts, function (part) {
        parent = parent[part] = parent[part] || {};
    });

    return parent = _.extend(parent, _.isFunction(obj) ? obj(jQuery) : obj);
};

namespace("FR.Util", function ($) {
    var self = {};

    self.queryString = function (param, qs) {
        var params = {};
        if (qs) {
            qs = _.str.contains(qs, "?") ? _.str.strRight("?") : "";
        }
        else {
            qs = _.str.strRight(location.search, "?");
        }

        _.each(qs.split('&'), function (q) {
            if (q) {
                var kvp = q.split('=');
                params[kvp[0].toString()] = decodeURIComponent(kvp[1].toString());
            }
        });
        return param ? params[param] : params;
    };

    self.joinQueryString = function (params) {
        return _.map(_.keys(params), function (k) { return k + "=" + encodeURIComponent(params[k]) }).join("&");
    };

    self.debug = function (arg) {
        var d = new Date();
        var dt = d.getHours() + ":" + d.getMinutes() + ":" + d.getSeconds() + "." + d.getMilliseconds();
        if (window.console && typeof window.console.log === "function") {
            window.console.log(dt + " > " + arg);
        }
    };

    var _datePatterns = {
        UKDate: /^(\d{2})\/(\d{2})\/(\d{4})$/,
        TZ: /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$/
    };

    self.ParseDate = function (s) {
        if (!s) { return null; }
        if (_.isDate(s)) return s;

        s = _.str.trim(s);
        if (_datePatterns.TZ.test(s)) {
            var arr = _datePatterns.TZ.exec(s);
            return new Date(arr[1], parseInt(arr[2]) - 1, arr[3], arr[4], arr[5], arr[6]);
        }
        if (_datePatterns.UKDate.test(s)) {
            var arr = _datePatterns.UKDate.exec(s);
            var dd = parseInt(arr[1]),
                mm = parseInt(arr[2]),
                yyyy = parseInt(arr[3]);
            if (mm <= 12 && dd <= 31) { //extra check to ensure it's not a mm/dd/yyyy
                return new Date(yyyy, mm - 1, dd);
            }
        }
        //add more cases as required
    };

    self.FormatDate = function (dt, format) {
        dt = self.ParseDate(dt); //ensure input is a correct date
        if (!dt || !_.isDate(dt)) return "";

        if (dt.localeFormat) { //Rely on this function added to the Date.prototype by msajaxbundle.js
            if (format == "standard") {
                // Tweak to avoid pre-daylight saving time dates to come as 23:00 of the day before
                dt.setHours(dt.getHours() + 12);
                return dt.localeFormat("dd MMM yyyy");
            } else if (format == "standardWithTime") {
                var meridiem = dt.getHours() < 12 ? 'AM' : 'PM';
                return dt.localeFormat('dd MMM yyyy hh:mm:ss') + ' ' + meridiem;
            }
        }

        return dt.toString();
    };

    self.GetImageRendition = function (imgUrl, renditionId) {
        if (imgUrl && imgUrl.toString) imgUrl = imgUrl.toString(); //To allow passing in Srch.ValueInfo objects as the imgUrl parameter
        if (!imgUrl) return "";
        return _.str.trim(_.str.strLeft(imgUrl, "?")) + (renditionId ? "?RenditionID=" + renditionId : "");
    };

    //Wraps a promise callback in another promise function so that multiple calls to the wrapped will only result 
    //in a single call of the original one, the rest being enqueued and called when it is resolved
    //Example: var ensureInfoLoaded = FR.Util.SingleCallPromise(loadInfo); ensureInfoLoaded().then(...);
    //In the example above, multiple calls to ensureInfoLoaded() will result in a single call to loadInfo()

    self.SingleCallPromise = function (promiseCallback) {
        var dfd = null;
        
        return function () {
            if (!dfd) {
                dfd = $.Deferred();
                promiseCallback().then(function (result) { dfd.resolve(result); }, function () { dfd.reject() });
            }
            return dfd.promise();
        };
    };

    return self;
});

namespace("FR.Util.SP", function ($) {
    var self = {};

    self.post = function (url, isVerboseElseNometadata) {
        var odataMode = isVerboseElseNometadata ? "verbose": "nometadata";
        return $.ajax({
            type: "POST",
            url: url,
            contentType: "application/json;odata=" + odataMode,
            headers: {
                "accept": "application/json;odata=" + odataMode,
                "X-RequestDigest": $("#__REQUESTDIGEST").val()
            }
        });
    };

    self.get = function (url, isVerboseElseNometadata) {
        var odataMode = isVerboseElseNometadata ? "verbose" : "nometadata";
        return $.ajax({
            url: url,
            contentType: "application/json;odata=" + odataMode,
            headers: {
                "accept": "application/json;odata=" + odataMode
            }
        });
    };

    self.deferredQuery = function (ctx) {
        var dfd = $.Deferred();

        ctx.executeQueryAsync(function () { dfd.resolve() }, function () { dfd.reject() });

        return dfd.promise();
    };

    //Usage 1: FR.Util.SP.ensureScriptsLoaded(["SP.js", ...], function() { ... });
    //Usage 2: $.when(FR.Util.SP.ensureScriptsLoaded(["SP.js", ...])).done(function() { ... });
    self.ensureScriptsLoaded = function (scripts, callback) {
        if (!_.isArray(scripts)) scripts = [scripts];

        var dfd = null;

        //If no callback specified, use as deferred
        if (!callback) {
            dfd = $.Deferred();
            callback = function () { dfd.resolve() };
        }

        //Ensure requested scripts are registered in SOD and loaded
        var wrappedCallbacks = _.reduce(scripts.reverse(), function (f, scriptName) {
            return function () {
                if (typeof _v_dictSod == "object") {
                    if (!_v_dictSod[scriptName]) { //If not registered > register and force load
                        SP.SOD.registerSod(scriptName, "/_layouts/15/" + scriptName);
                        SP.SOD.executeFunc(scriptName, null, function () { });
                    }
                    else if (_v_dictSod[scriptName].state != Sods.loaded) { //Registered but not loaded
                        SP.SOD.executeFunc(scriptName, null, function () { });
                    }
                }
                SP.SOD.executeOrDelayUntilScriptLoaded(f, scriptName);
            };
        }, callback);

        wrappedCallbacks();

        return dfd ? dfd.promise() : undefined;
    };

    self.getCurrentUserInfo = function () {
        var userid = _spPageContextInfo.userId;
        var url = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid(" + userid + ")";
        return self.get(url);
    };

    //Takes a tag and an array of values and returns an XML string of the balanced binary tree
    var _binaryOp = function (tag, arr) {
        if (!_.isArray(arr) || arr.length == 0) return null;
        if (arr.length == 1) return arr[0];

        var cut = Math.ceil(arr.length / 2),
            arrLeft = _.first(arr, cut),
            arrRight = _.rest(arr, cut),
            left = _binaryOp(tag, arrLeft),
            right = _binaryOp(tag, arrRight);

        if (left && right) { return "<" + tag + ">" + left + right + "</" + tag + ">"; }
        if (!left) return right;
        return left;
    };

    self.CAML = function () {
        return {
            Condition: function (condition, fieldName, fieldValue, fieldType) {
                return "<" + condition + "><FieldRef Name='" + fieldName + "'/><Value Type='" + (fieldType || "Text") + "'>" + fieldValue + "</Value></" + condition + ">";
            },
            Contains: function (fieldName, fieldValue, fieldType) {
                return self.CAML.Condition("Contains", fieldName, fieldValue, fieldType);
            },
            IsNull: function (fieldName) {
                return "<IsNull><FieldRef Name='" + fieldName + "'/></IsNull>";
            },
            IsNotNull: function (fieldName) {
                return "<IsNotNull><FieldRef Name='" + fieldName + "'/></IsNotNull>";
            },
            FilterByID: function (id) {
                return self.CAML.Condition("Eq", "ID", id, "Counter");
            },
            FilterByLookupId: function (fieldName, id) {
                return _.str.sprintf("<Eq><FieldRef Name='%s' LookupId='true' /><Value Type='Lookup'>%d</Value></Eq>", fieldName, id);
            },
            And: function () {
                return _binaryOp("And", _.flatten(arguments));
            },
            Or: function () {
                return _binaryOp("Or", _.flatten(arguments));
            },
            WrapViewFields: function (fields) {
                return "<ViewFields>" + _.map(fields, function (e) { return "<FieldRef Name='" + e + "'/>" }).join("") + "</ViewFields>";
            },
            WrapOrderBy: function (field, orderDesc) {
                var s = "";
                if (_.isArray(field)) {
                    s = _.map(field, function (f) { return "<FieldRef Name='" + f.field + "'" + (f.orderDesc ? " Ascending='FALSE'" : "") + "/>"; }).join("");
                }
                else {
                    s = "<FieldRef Name='" + field + "'" + (orderDesc ? " Ascending='FALSE'" : "") + "/>";
                }
                return "<OrderBy>" + s + "</OrderBy>";
            }
        };
    }();

    self.replaceSPTokens = function (s, isCustomToken, isServerRelative) {
        var tokenPrefix = isCustomToken ? "#" : "~";
        var siteCollUrl = _.str.rtrim(isServerRelative ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.siteAbsoluteUrl, "/");
        var siteUrl = _.str.rtrim(isServerRelative ? _spPageContextInfo.webServerRelativeUrl : _spPageContextInfo.webAbsoluteUrl, "/");
        var siteCollRegex = new RegExp(tokenPrefix + "sitecollection", "gi");
        var siteRegex = new RegExp(tokenPrefix + "site", "gi");

        if (siteCollUrl || siteCollUrl === "") {
            s = s.replace(siteCollRegex, siteCollUrl);
        }
        if (siteUrl || siteUrl === "") {
            s = s.replace(siteRegex, siteUrl);
        }
        return s;
    };

    //Extracts data from either a SP.Taxonomy.TaxonomyFieldValue or SP.Taxonomy.TaxonomyFieldValueCollection object
    //and returns it in an array of { id: "", label: "" } objects
    self.parseTaxonomyFieldValuesAsArray = function (o) {
        var _parseTaxonomyFieldValue = function (tfv) {
            return {
                id: tfv.get_termGuid(),
                label: tfv.get_label()
            };
        }

        if (!o) return [];
        if (o.constructor == SP.Taxonomy.TaxonomyFieldValue) {
            return [_parseTaxonomyFieldValue(o)];
        }
        else if (o.constructor == SP.Taxonomy.TaxonomyFieldValueCollection) {
            return _.times(o.get_count(), function (i) {
                return _parseTaxonomyFieldValue(o.getItemAtIndex(i));
            });
        }
    };

    var _sewpTemplate =
        '<webParts><webPart xmlns="http://schemas.microsoft.com/WebPart/v3">' +
         '<metaData>' +
          '<type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" />' +
         '</metaData>' +
         '<data><properties>' +
          '<property name="Title" type="string">#title#</property>' +
          '<property name="ChromeType" type="chrometype">None</property>' +
          '<property name="Content" type="string"><![CDATA[#content#]]></property>' +
         '</properties></data>' +
        '</webPart></webParts>';

    self.GetSEWPMarkup = function (title, content) {
        return _sewpTemplate.replace(/#title#/gi, title).replace(/#content#/gi, content);
    };

    self.RedirectTo = function (urlToRedirect) {
        var qs = FR.Util.queryString();
        if (qs["ControlMode"] == "Edit" || qs["DisplayMode"] == "Design") return; //Avoid redirections while editing pages

        var url = self.replaceSPTokens(urlToRedirect || "~sitecollection");
        location.href = url;
    };



    return self;
});

namespace("FR.Util.UI", function ($) {
    var self = {};

    self.AddStatus = function (message, colour) {
        var statusId = SP.UI.Status.addStatus(message);
        if (colour !== undefined && colour !== null) {
            SP.UI.Status.setStatusPriColor(statusId, colour);
        }
    }

    self.RemoveStatus = function (statusId) {
        SP.UI.Status.removeStatus(statusId);
    }

    self.RemoveAllStatus = function () {
        SP.UI.Status.removeAllStatus(true);
    }

    return self;
});

namespace("FR.Util.DisplayTemplates", function ($) {
    var self = {};

    var _resultTypesByFileExtension = {
        "Word": ["doc", "docx", "dot", "dotx"],
        "PowerPoint": ["ppt", "pptx"],
        "Excel": ["xls", "xlsx"],
        "OneNote": ["one"],
        "Picture": ["png", "jpg", "jpeg", "gif"],
        "WebPage": ["aspx"],//, "html", "htm"],
        //Add more as required
        "PDF": ["pdf"]
    };
    var _defaultHoverPanel = "PDF";

    self.GetHoverPanelTemplateUrl = function (fileType) {
        var templateName = _.find(_.keys(_resultTypesByFileExtension), function (k) {
            return _.contains(_resultTypesByFileExtension[k], fileType);
        }) || _defaultHoverPanel;
        return "~sitecollection/_catalogs/masterpage/Display Templates/Search/Item_" + templateName + "_HoverPanel.js";
    };

    return self;
});

namespace("FR.Util.Taxonomy", function ($) {
    var self = {};

    var getTree = function (terms, termsEnumerator, selectedId) {
        var tree = {
            term: terms,
            children: []
        };

        // Loop through each term
        while (termsEnumerator.moveNext()) {
            var currentTerm = termsEnumerator.get_current();
            var currentTermPath = currentTerm.get_pathOfTerm().split(';');
            var children = tree.children;

            // Loop through each part of the path
            for (var i = 0; i < currentTermPath.length; i++) {
                var foundNode = false;

                for (var j = 0; j < children.length; j++) {
                    if (children[j].name === currentTermPath[i]) {
                        foundNode = true;
                        break;
                    }
                }

                // Select the node, otherwise create a new one
                var term = foundNode ? children[j] : { name: currentTermPath[i], children: [] };

                // If we're a child element, add the term properties
                if (i === currentTermPath.length - 1) {
                    term.term = currentTerm;
                    term.title = currentTerm.get_name();
                    term.guid = currentTerm.get_id().toString();
                    term.description = currentTerm.get_description();
                    term.localCustomProperties = currentTerm.get_localCustomProperties();
                    if (selectedId !== undefined && selectedId !== null && selectedId === term.guid) {
                        term.selected = true;
                    } else {
                        term.selected = false;
                    }
                }

                // If the node did exist, let's look there next iteration
                if (foundNode) {
                    children = term.children;
                }
                    // If the segment of path does not exist, create it
                else {
                    children.push(term);

                    // Reset the children pointer to add there next iteration
                    if (i !== currentTermPath.length - 1) {
                        children = term.children;
                    }
                }
            }
        }
        return tree;
    }

    // ==========================================================================================
    // @param {string} id - Termset ID
    // @param {object} callback - Callback function to call upon completion and pass termset into
    // ------------------------------------------------------------------------------------------
    // Returns a termset, based on ID
    // ==========================================================================================
    self.getTermSet = function (id, callback) {
        SP.SOD.loadMultiple(['sp.js'], function () {
            // Make sure taxonomy library is registered
            SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));

            SP.SOD.loadMultiple(['sp.taxonomy.js'], function () {
                var ctx = SP.ClientContext.get_current();
                var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);
                var termStore = taxonomySession.getDefaultSiteCollectionTermStore();
                var termSet = termStore.getTermSet(id);
                var terms = termSet.getAllTerms();

                ctx.load(terms);
                ctx.executeQueryAsync(
                    Function.createDelegate(this, function (sender, args) {
                        callback(terms);
                    }),
                    Function.createDelegate(this, function (sender, args) {
                        if (typeof window.console !== 'undefined' && typeof window.console.log !== 'undefined') {
                            console.log('Error: FR.Util.Taxonomy.getTermSet() Cannot invoke getAllTerms() method or retrieve property for ' + id);
                        }
                    }));
            });
        });
    };

    // ==========================================================================================
    // @param {obj} tree The term tree
    // @return {obj} A sorted term tree
    // ------------------------------------------------------------------------------------------
    // Returns sort children array of a term tree by a sort order
    // ==========================================================================================
    self.sortTermsFromTree = function (tree) {
        // Check to see if the get_customSortOrder function is defined. If the term is actually a term collection,
        // there is nothing to sort.
        if (tree.children.length && tree.term.get_customSortOrder) {
            var sortOrder = null;

            if (tree.term.get_customSortOrder()) {
                sortOrder = tree.term.get_customSortOrder();
            }

            // If not null, the custom sort order is a string of GUIDs, delimited by a :
            if (sortOrder) {
                sortOrder = sortOrder.split(':');
                tree.children.sort(function (a, b) {
                    var indexA = sortOrder.indexOf(a.guid);
                    var indexB = sortOrder.indexOf(b.guid);
                    if (indexA > indexB) {
                        return 1;
                    } else if (indexA < indexB) {
                        return -1;
                    }
                    return 0;
                });
            }
                // If null, terms are just sorted alphabetically
            else {
                tree.children.sort(function (a, b) {
                    if (a.title > b.title) {
                        return 1;
                    } else if (a.title < b.title) {
                        return -1;
                    }
                    return 0;
                });
            }
        }

        for (var i = 0; i < tree.children.length; i++) {
            tree.children[i] = FR.Util.Taxonomy.sortTermsFromTree(tree.children[i]);
        }

        return tree;
    };

    // ==========================================================================================
    // @param {string} id - Termset ID
    // @param {object} callback - Callback function to call upon completion and pass termset into
    // ------------------------------------------------------------------------------------------
    // Returns an array object of terms as a tree
    // ==========================================================================================
    self.getTermSet_callback = function (id, selectedId, callback) {
        FR.Util.Taxonomy.getTermSet(id, function (terms) {
            var termsEnumerator = terms.getEnumerator();
            var tree = getTree(terms, termsEnumerator, selectedId);

            tree = FR.Util.Taxonomy.sortTermsFromTree(tree);
            callback(tree);
        });
    };

    // ==========================================================================================
    // @param {string} id - Termset ID
    // @param {object} callback - Callback function to call upon completion and pass termset into
    // ------------------------------------------------------------------------------------------
    // Returns an array object of terms as a tree
    // ==========================================================================================
    self.getTermSet_deferred = function (id, selectedId) {
        var dfd = jQuery.Deferred();

        FR.Util.Taxonomy.getTermSet(id, function (terms) {
            var termsEnumerator = terms.getEnumerator();
            var tree = getTree(terms, termsEnumerator, selectedId);

            tree = FR.Util.Taxonomy.sortTermsFromTree(tree);
            dfd.resolve(tree);
        });

        return dfd.promise();
    };

    return self;
});