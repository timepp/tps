/*
    https://github.com/timepp/tps
    2013.1.4

    TPS is a javascript library for local scripts(.wsf, .js, .hta).
*/

if (!this.tps) {
    this.tps = {};
}

// well-known windows object
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");
var net = new ActiveXObject("WScript.Network");
var shellapp = new ActiveXObject("Shell.Application");
var env = shell.Environment("Process");

var HKEY_LOCAL_MACHINE = 0x80000002;
var HKEY_CURRENT_USER = 0x80000001;
var HKEY_CLASSES_ROOT = 0x80000000;

// add trim function to String
String.prototype.trim = function () {
    return this.replace(/^\s+|\s+$/g, "");
};
String.prototype.ltrim = function () {
    return this.replace(/^\s+/, "");
};
String.prototype.rtrim = function () {
    return this.replace(/\s+$/, "");
};
String.prototype.lpad = function (ch, l) {
    var ret = this;
    while (ret.length < l) ret = ch + ret;
    return ret;
};
String.prototype.rpad = function (ch, l) {
    var ret = this;
    while (ret.length < l) ret += ch;
    return ret;
};
String.prototype.icaseEqual = function (str) {
    return this.toLowerCase() == str.toLowerCase();
};
if (typeof String.prototype.format !== 'function') {
    String.prototype.format = function () {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function (match, number) {
            return typeof args[number] != 'undefined'
              ? args[number]
              : match
            ;
        });
    };
}
if (typeof String.prototype.beginWithOneOf !== 'function') {
    String.prototype.beginWithOneOf = function (arr) {
        for (var i in arr) {
            if (this.toLowerCase().indexOf(arr[i].toLowerCase()) >= 0) return true;
        }
        return false;
    };
}
if (typeof String.prototype.startsWith !== 'function') {
    String.prototype.startsWith = function (b) {
        var i = b.length;
        if (this.length < i) {
            return false;
        }

        while (i--) {
            if (this.charAt(i) != b[i]) {
                return false;
            }
        }

        return true;
    };
}
if (typeof String.prototype.endsWith !== 'function') {
    String.prototype.endsWith = function (suffix) {
        return this.indexOf(suffix, this.length - suffix.length) !== -1;
    };
}


(function () {

    var ForReading = 1, ForWriting = 2;

    // tps.util ============================================================================================================================
    // tps.util ============================================================================================================================
    // tps.util ============================================================================================================================
    // tps.util ============================================================================================================================
    // tps.util ============================================================================================================================

    tps.util = {
        SingleQuote: function (str) {
            return "'" + str + "'";
        },
        DoubleQuote: function (str) {
            return '"' + str + '"';
        },
        RemoveQuote: function (str) {
            return str.replace(/^(["'])(.*)\1$/, "$2");
        },
        MergeProperty: function (o, p) {
            for (var key in p) {
                o[key] = p[key];
            }
        },
        SubObject: function () {
            var obj = arguments[0];
            for (var i = 1; i < arguments.length; i++) {
                var arg = arguments[i];
                var pn = arg.replace(/^\[(.*)\]$/, "$1");
                if (!(pn in obj)) obj[pn] = (pn == arg ? new Object : new Array);
                obj = obj[pn];
            }
            return obj;
        },
        IndexOf: function (arr, obj, fromIndex) {
            if (fromIndex == null) {
                fromIndex = 0;
            } else if (fromIndex < 0) {
                fromIndex = Math.max(0, arr.length + fromIndex);
            }
            for (var i = fromIndex, j = arr.length; i < j; i++) {
                if (arr[i] === obj)
                    return i;
            }
            return -1;
        },
        // group array to map by some properties
        // arr: Array of Objects
        // hash_function: map function, translate Object to String
        // -> {"hash":[object_has_same_hash]}
        Group: function (arr, hash_function) {
            var g = {};
            for (var index in arr) {
                var obj = arr[index];
                var hash = hash_function(obj);
                tps.util.SubObject(g, "[" + hash + "]").push(obj);
            }
            return g;
        },
        // accumulate by same property of some objects
        Accumulate: function (arr, prop) {
            var sum = 0;
            for (var index in arr) {
                sum += arr[index][prop];
            }
            return sum;
        },
        FormatDateString: function (dt, fmt) {
            var ret = "";
            var D2 = function (n) { return n.toString().lpad("0", 2); };
            for (var i = 0; i < fmt.length; i++) {
                var c = fmt.charAt(i);
                if (c == "Y") ret += dt.getFullYear();
                else if (c == "m") ret += D2(dt.getMonth() + 1);
                else if (c == "d") ret += D2(dt.getDate());
                else if (c == "H") ret += D2(dt.getHours());
                else if (c == "M") ret += D2(dt.getMinutes());
                else if (c == "S") ret += D2(dt.getSeconds());
                else if (c == "I") ret += dt.getMilliseconds().toString().lpad('0', 3);
                else ret += c;
            }
            return ret;
        },
        ParseDateString: function (datestr) {
            /*
                absolute time formats: now | today | yyyymmdd | yyyymmddHHMMSS
                relative time formats: <absolute time format>[+-]number[ymdHMS]
            */
            function ParseAbsDateString(datestr) {
                if (datestr == "now") {
                    return new Date();
                }
                else if (datestr == "today") {
                    var dt = new Date();
                    return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), 0, 0, 0, 0);
                }
                else {
                    var re_abs = /^(\d\d\d\d)(\d\d)(\d\d)((\d\d)(\d\d)(\d\d))?$/;
                    var m = re_abs.exec(datestr);
                    if (m) {
                        return new Date(m[1], m[2] - 1, m[3], m[5]?m[5]:0, m[6]?m[6]:0, m[7]?m[7]:0, 0);
                    }
                }
                return new Date();
            }

            var re_rela = /^(.*)([+-])([0-9]+)([ymdHMS])$/;
            var m = re_rela.exec(datestr);
            if (m) {
                var dtbase = ParseAbsDateString(m[1]);
                var factor = (m[2] == "+" ? 1 : -1);
                var n = parseInt(m[3]);

                if (m[4] == "y" || m[4] == 'm') {
                    var dy = dtbase.getFullYear();
                    var dm = dtbase.getMonth();
                    var dd = dtbase.getDate();
                    var dH = dtbase.getHours();
                    var dM = dtbase.getMinutes();
                    var dS = dtbase.getSeconds();
                    if (m[4] == "y") dy += factor * n;
                    else if (m[4] == "m") {
                        dm += factor * n;
                        dy += Math.floor(dm / 12);
                        dm = dm % 12;
                    }
                    return new Date(dy, dm, dd, dH, dM, dS, 0);
                }
                else {
                    var ms = dtbase.getTime();
                    switch (m[4]) {
                        case "S": n *= 1; break;
                        case "M": n *= 60; break;
                        case "H": n *= 3600; break;
                        case "d": n *= 86400; break;
                    }
                    return new Date(ms + n * factor * 1000);
                }
            }
            else {
                return ParseAbsDateString(datestr);
            }
        },
        SplitCmdLine: function (cmdline) {
            args = [];
            var re = /^\s*("[^"]+"|[^" ]+)/;
            while ((arr = re.exec(cmdline)) != null) {
                args.push(tps.util.RemoveQuote(arr[1]));
                cmdline = cmdline.substr(arr.lastIndex);
            }
            return args;
        },

        /* 参数形式：
           --opt=val         设置选项opt值为val
           -switch           打开开关switch
           target            添加一个target
           --                特殊占位符，表明其后的所有参数都做为target
        */
        ParseArgument: function (argv) {
            var args = { _targets: [] };
            for (var i = 0; i < argv.length; i++) {
                var str = argv[i];
                if (str == "--") {
                    args._targets.concat(argv.slice(i + 1));
                    break;
                } else if (str.substring(0, 2) == "--") {
                    var j = str.indexOf("=");
                    args[str.slice(2, j)] = tps.util.RemoveQuote(str.slice(j + 1));
                } else if (str.charAt(0) == "-") {
                    args[str.slice(1)] = true;
                } else {
                    args._targets.push(str);
                }
            }
            return args;
        },

        CreateGUID: function () {
            return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g,
                function (c) {
                    var r = Math.random() * 16 | 0, v = c == 'x' ? r : r & 0x3 | 0x8;
                    return v.toString(16);
                }
            )
        },
        arraysEqual: function (arr1, arr2) {
            if (arr1.length !== arr2.length) return false;
            for (var i = arr1.length; i--;) {
                if (arr1[i] !== arr2[i]) return false;
            }
            return true;
        }
    };

    tps.sys = {
        IsAmd64: function () {
            var path = shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%");
            return fso.FolderExists(path);
        },
        HasFullPrivilege: function () {
            try {
                var value = shell.RegRead("HKEY_USERS\\S-1-5-19\\");
            } catch (e) {
                return false;
            }
            return true;
        },
        RestartHTA: function (requestAdmin, escapeWOW64) {
            var mshta = "mshta.exe";
            var verb = "open";
            var needRestart = false;
            if (escapeWOW64) {
                var sysnativePath = shell.ExpandEnvironmentStrings("%windir%\\sysnative");
                if (fso.FolderExists(sysnativePath)) {
                    mshta = sysnativePath + "\\mshta.exe";
                    needRestart = true;
                }
            }
            if (requestAdmin) {
                if (!tps.sys.HasFullPrivilege()) {
                    verb = "runas";
                    mshta = "mshta.exe";
                    needRestart = true;
                }
            }
            if (needRestart) {
                shellapp.ShellExecute(mshta, tps.sys.GetScriptPath(), "", verb, 1);
                window.close();
                body.onload = null;
                return true;
            }
            return false;
        },
        GetScriptPath: function () {
            try {
                return WScript.ScriptFullName;
            } catch (e) {
                // for some reason bootstrap will add '/' at the begin of pathname
                var pathname = document.location.pathname.replace(/^\/(.*)$/, "$1");
                if (document.location.hostname) {
                    pathname = "\\\\" + document.location.hostname + pathname;
                }
                return pathname;
            }
        },
        GetScriptDir: function () {
            return tps.file.GetDir(tps.sys.GetScriptPath());
        },
        NotifySettingChange: function(name) {
            return shell.Run("calldll.exe user32.dll SendNotifyMessageW int:0xffff int:0x1A int:0 wstr:" + name, 0, true);
        },

        systemEnvRegPath: "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment",
        userEnvRegPath: "HKCU\\Environment",
        // Note: It seems that the original type of 'path' is "REG_EXPAND_SZ", but some software may change it to "REG_SZ"
        //       We need to handle type when read, and use official one when save
        GetSystemEnv: function (vname) {
            try {
                return tps.reg.GetExpandStringValue(tps.sys.systemEnvRegPath, vname);
            } catch (e) {
                return tps.reg.GetStringValue(tps.sys.systemEnvRegPath, vname);
            }
        },
        SetSystemEnv: function (vname, val) {
            tps.reg.SetExpandStringValue(tps.sys.systemEnvRegPath, vname, val);
            tps.sys.NotifySettingChange("Environment");
        },
        DeleteSystemEnv: function (vname) {
            tps.reg.DeleteValue(tps.sys.systemEnvRegPath, vname);
        },
        GetUserEnv: function (vname) {
            try {
                return tps.reg.GetExpandStringValue(tps.sys.userEnvRegPath, vname);
            } catch (e) {
                return tps.reg.GetStringValue(tps.sys.userEnvRegPath, vname);
            }
        },
        SetUserEnv: function (vname, val) {
            tps.reg.SetExpandStringValue(tps.sys.userEnvRegPath, vname, val);
            tps.sys.NotifySettingChange("Environment");
        },
        DeleteUserEnv: function (vname) {
            tps.reg.DeleteValue(tps.sys.userEnvRegPath, vname);
        },
        
        SetEnv: function (name, val) {
            if (val == null) {
                env.Remove(name);
            }
            else {
                env.Item(name) = val;
            }
        },
        GetEnv: function (name) {
            return env.Item(name);
        },
        InPath: function (path) {
            var paths = tps.sys.GetSystemEnv("path").toLowerCase().split(";");
            return tps.util.IndexOf(paths, path.toLowerCase()) != -1;
        },
        AddToPath: function (path) {
            if (!tps.sys.InPath(path)) {
                var pathval = tps.sys.GetSystemEnv("path");
                var newpathval = pathval;
                if (newpathval.slice(-1) != ";") newpathval += ";";
                newpathval += path;
                tps.sys.SetSystemEnv("path", newpathval);
            }
        },
        RunCommandAndGetResult: function (cmdline, of, ef) {
            var outfile = of ? of : shell.ExpandEnvironmentStrings("%temp%") + "\\" + fso.GetTempName();
            var errfile = ef ? ef : shell.ExpandEnvironmentStrings("%temp%") + "\\" + fso.GetTempName();

            try {
                fso.DeleteFile(outfile);
                fso.DeleteFile(errfile);
            } catch (e) { }

            cmdline = "cmd.exe /C " + cmdline + ' > "OUT" 2> "ERR"'.replace("OUT", outfile).replace("ERR", errfile);
            tps.log.Debug("RUN:[" + cmdline + "]");
            var returnValue = shell.Run(cmdline, 0, true);
            var ret = {
                output: tps.file.ReadTextFileSimple(outfile),
                errors: tps.file.ReadTextFileSimple(errfile),
                retval: returnValue
            };

            try {
                if (!of) fso.DeleteFile(outfile);
                if (!ef) fso.DeleteFile(errfile);
            } catch (e) {}

            return ret;
        }
    };

    tps.reg = {
        GetGeneralValueAsString: function (key, valname, matcher) {
            var vpart = valname ? "/v " + tps.util.DoubleQuote(valname) : '/ve';
            var cmdline = 'reg query "{0}" {1}'.format(key, vpart);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval) throw result.errors;
            var m = matcher.exec(result.output);
            if (!m || !m[1]) {
                throw "reg output parse failed";
            }
            return m[1];
        },
        SetGeneralValueByString: function (key, valname, type, val) {
            var vpart = valname ? "/v " + tps.util.DoubleQuote(valname) : '/ve';
            // escape all " in val, see http://stackoverflow.com/questions/562038/escaping-double-quotes-in-batch-script
            if (type == "REG_SZ" || type == "REG_MULTI_SZ" || type == "REG_EXPAND_SZ") {
                val = val.replace(/"/g, "\\\"");
            }
            var cmdline = 'reg add "{0}" {1} /t {2} /d "{3}" /f'.format(key, vpart, type, val);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval) throw result.errors;
        },
        GetStringValue: function (key, valname) {
            return tps.reg.GetGeneralValueAsString(key, valname, /REG_SZ\s+([^\r\n]*)/gm);
        },
        GetExpandStringValue: function (key, valname) {
            return tps.reg.GetGeneralValueAsString(key, valname, /REG_EXPAND_SZ\s+([^\r\n]*)/gm);
        },
        GetMultiStringValue: function (key, valname) {
            return tps.reg.GetGeneralValueAsString(key, valname, /REG_MULTI_SZ\s+([^\r\n]*)/gm).split("\\0");
        },
        GetIntValue: function (key, valname) {
            return parseInt(tps.reg.GetGeneralValueAsString(key, valname, /REG_DWORD\s+0[xX](\S+)/gm), 16);
        },
        GetBoolValue: function (key, valname) {
            return GetIntValue(key, valname) > 0;
        },
        SetStringValue: function (key, valname, val) {
            tps.reg.SetGeneralValueByString(key, valname, "REG_SZ", val);
        },
        SetExpandStringValue: function (key, valname, val) {
            tps.reg.SetGeneralValueByString(key, valname, "REG_EXPAND_SZ", val);
        },
        SetMultiStringValue: function (key, valname, val) {
            tps.reg.SetGeneralValueByString(key, valname, "REG_MULTI_SZ", val.join("\\0"));
        },
        SetIntValue: function (key, valname, val) {
            tps.reg.SetGeneralValueByString(key, valname, "REG_DWORD", val);
        },
        SetBoolValue: function (key, valname, val) {
            tps.reg.SetIntValue(key, valname, val ? 1 : 0);
        },
        StringValueExists: function (key, valname) {
            try {
                tps.reg.GetStringValue(key, valname);
            } catch (e) {
                return false;
            }
            return true;
        },
        IntValueExists: function (key, valname) {
            try {
                tps.reg.GetIntValue(key, valname);
            } catch (e) {
                return false;
            }
            return true;
        },
        CreateKey: function (key) {
            var cmdline = 'reg add "{0}" /f'.format(key);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval) throw result.errors;
        },
        DeleteKey: function (key) {
            var cmdline = 'reg delete "{0}" /f'.format(key);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval) throw result.errors;
        },
        DeleteValue: function (key, val) {
            var cmdline = 'reg delete "{0}" /v "{1}" /f'.format(key, val);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval) throw result.errors;
        },
        KeyExisis: function (key) {
            var cmdline = 'reg query "{0}" /ve'.format(key);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            return result.retval == 0;
        },
        EnumSubKeys: function (key) {
            var cmdline = 'reg query "{0}" /f * /k'.format(key);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval != 0) throw result.errors;
            var lines = result.output.split("\n");
            var subkeys = [];
            for (var i in lines) {
                if (lines[i].startsWith("HKEY_")) {
                    subkeys.push(lines[i].trim());
                }
            }
            return subkeys;
        },
        // returns [{valuename, valuetype, valuestr}]
        EnumValues: function (key) {
            var cmdline = 'reg query "{0}" /f * /v'.format(key);
            var result = tps.sys.RunCommandAndGetResult(cmdline);
            if (result.retval != 0) throw result.errors;
            var lines = result.output.split("\n");
            var values = [];
            for (var i in lines) {
                var m = /^(.*?)\s+(REG_\S+)\s+(.*)$/.exec(lines[i].trim());
                if (m && m[1] != "(Default)")
                    values.push({ name: m[1], type: m[2], valstr: m[3] });
            }
            return values;
        },
        BatchGetStringValues: function (key, subkeys, valname) {
            var keys = tps.reg.EnumKeys(key);
            var values = [];
            if (subkeys == null) subkeys = ""; else subkeys = subkeys + "\\";
            for (var i = 0; i < keys.length; i++) {
                var val = tps.reg.GetStringValue(key + "\\" + keys[i] + "\\" + subkeys, valname);
                values.push(val);
            }
            return values;
        },
        OpenRegEdit: function (path) {
            tps.reg.SetStringValue("HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit", "LastKey", path);
            shell.Run("regedit.exe");
        }
    };

    tps.ui = {
        ClearTable: function (tbl) {
            while (tbl.rows.length > 0) {
                tbl.deleteRow(0);
            }
        },
        SelectItem: function (select, text, defaultIndex) {
            if (!defaultIndex) defaultIndex = 0;
            for (var i = 0; i < select.length; i++) {
                var opt = select.item(i);
                if (opt.value == text) {
                    opt.selected = true;
                    return;
                }
            }
            var opt = select.item(defaultIndex);
            if (opt) opt.selected = true;
        },
        AddGradientBK: function (o, c1, c2, tp) {
            if (c2 == null) c2 = "#FFFFFF";
            if (tp == null) tp = 1;
            o.style.filter = "progid:DXImageTransform.Microsoft.Gradient(GradientType=%tp,StartColorStr='%sc', EndColorStr='%ec')".replace("%sc", c1).replace("%ec", c2).replace("%tp", tp);
        },
        CenterAbsoluteObject: function (o) {
            o.style.left = (document.body.clientWidth - o.offsetWidth) / 2;
            o.style.top = (document.body.clientHeight - o.offsetHeight) / 2;
        },
        PopupMessage: function (title, msg) {
            var oDiv = document.createElement("div");
            oDiv.id = "div_err";
            oDiv.style.cssText = "padding: 5px; border: medium #0000FF double; position: absolute; font-family: monospace; background-color:#FFFFFF";

            var oSpan = document.createElement("span");
            oSpan.innerHTML = title + "<br/>";
            oSpan.style.fontWeight = "bold";
            oDiv.appendChild(oSpan);

            oSpan = document.createElement("span");
            oSpan.innerText = msg;
            oSpan.style.fontSize = "small";
            oDiv.appendChild(oSpan);

            oSpan = document.createElement("div");
            oSpan.align = "center";
            oSpan.innerHTML = "<button onclick=\"document.body.removeChild(document.getElementById('div_err'))\">Close</button>";
            oDiv.appendChild(oSpan);
            document.body.appendChild(oDiv);
            CenterAbsoluteObject(oDiv);
        },
        PlainTextToHTML: function (text) {
            var html = "";
            for (var i = 0; i < text.length; i++) {
                var code = text.charCodeAt(i);
                switch (code) {
                    case 38: html += "&amp;"; break;
                    case 60: html += "&lt;"; break;
                    case 62: html += "&gt;"; break;
                    case 10: html += "<br />"; break;
                    case 20: html += "&nbsp;"; break;
                    default: html += String.fromCharCode(code);
                }
            }
            return html;
        },
        HtmlCode_Span: function (text, className) {
            if (className)
                return '<span class="' + className + '">' + tps.ui.PlainTextToHTML(text) + '</span>';
            else
                return '<span>' + tps.ui.PlainTextToHTML(text) + '</span>';
        },
        DisableContextMenu: function (obj) {
            obj.oncontextmenu = function () { return false; };
        }
    };

    tps.file = {
        ReadTextFile: function (path, encoding) {
            var stream = new ActiveXObject('ADODB.Stream');
            stream.Type = 2;
            stream.Mode = 3;
            if (encoding) stream.Charset = encoding;
            stream.Open();
            stream.Position = 0;

            stream.LoadFromFile(path);
            var size = stream.Size;
            var text = stream.ReadText();

            stream.Close();

            return text;
        },
        WriteTextFile: function (text, path, encoding) {
            var stream = new ActiveXObject('ADODB.Stream');
            stream.Type = 2;
            stream.Mode = 3;
            if (encoding) stream.Charset = encoding;
            stream.Open();
            stream.Position = 0;

            stream.WriteText(text);
            try {
                fso.CreateFolder(tps.file.GetDir(path));
            } catch (e) { }
            stream.SaveToFile(path, 2);

            stream.Close();
        },
        CreateShortcut: function (lnkPath, targetPath, workingDir, argument, desc) {
            var oLnk = shell.CreateShortcut(lnkPath);
            oLnk.TargetPath = targetPath;
            oLnk.WorkingDirectory = workingDir;
            oLnk.Arguments = argument;
            oLnk.Description = desc;
            oLnk.Save();
        },
        ReadTextFileSimple: function (filename) {
            var content = "";
            try {
                var ofile = fso.OpenTextFile(filename, ForReading);
                content = ofile.ReadAll();
                ofile.Close();
            }
            catch (e) {
                //alert(e.message);
            }
            return content;
        },
        WriteTextFileSimple: function (text, filename) {
            try {
                fso.CreateFolder(tps.file.GetDir(filename));
            } catch (e) {
            }

            try {
                var ofile = fso.OpenTextFile(filename, ForWriting, true);
                ofile.Write(text);
                ofile.Close();
            }
            catch (e) {
                alert(e.message);
            }
        },
        ReadBinaryFile: function (filename, pos, len) {
            var stream = new ActiveXObject("ADODB.Stream");
            // stream.Type = 1;
            stream.Open();
            stream.LoadFromFile(filename);
            stream.Position = pos;
            var str = "";
            // str = stream.Read(len);
            str = stream.ReadText(len);
            stream.Close();
            return str;
        },
        // filter: regex for match file or dir
        // depth: 1 for immediately children, 0 for fully recursive
        // returns:
        // [ subdir1\, subdir1\file1, subdir1\file2, subdir1\subdir2\, subdir1\subdir2\file3, subdir3\file5, file6, ...]
        Glob: function (dir, filter, depth, arr, subdir) {
            if (!arr) arr = [];
            if (!subdir) subdir = "";
            if (depth == undefined) depth = 0;
            depth--;
            var d = fso.GetFolder(subdir ? dir + "\\" + subdir : dir);
            for (var fc = new Enumerator(d.files) ; !fc.atEnd() ; fc.moveNext()) {
                var f = fc.item();
                if (!filter || filter.test(f.Name)) {
                    arr.push(subdir + f.Name);
                }
            }
            for (var fc = new Enumerator(d.SubFolders) ; !fc.atEnd() ; fc.moveNext()) {
                var f = fc.item();
                //if (!filter || filter.test(f.Name)) {
                if (true) {
                    arr.push(subdir + f.Name + "\\");
                    if (depth != 0) {
                        tps.file.Glob(dir, filter, depth, arr, subdir + f.Name + "\\");
                    }
                }
            }
            return arr;
        },
        GetDir: function (path) {
            var dir = path.replace(/^(.*)[\\/][^\\/]+$/, "$1");
            if (dir == path) dir = ".";
            return dir;
        },
        GetFileName: function (path) {
            var fn = path.replace(/^.*[\\/]([^\\/]+)$/, "$1");
            return fn;
        },
        CreateFromTemplate: function (tfile, ofile, map, encoding) {
            if (!encoding) encoding = "UTF-8";

            var content = tps.file.ReadTextFile(tfile, encoding);
            for (var t in map) {
                var re = new RegExp(t, "g");
                content = content.replace(re, map[t]);
            }

            tps.file.WriteTextFile(content, ofile, encoding);
        },
        ZipDir: function (dir, path) {
            // write header
            tps.file.WriteTextFileSimple(
                "PK\x05\x06\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00",
                path);

            // copy
            var ns_dst = shellapp.NameSpace(fso.GetAbsolutePathName(path));
            var ns_src = shellapp.NameSpace(fso.GetAbsolutePathName(dir));
            ns_dst.CopyHere(ns_src.Items());

            // wait complete
            while (ns_dst.Items().count < ns_src.Items().count) {
                WScript.Sleep(1000);
            }
        },
        // Replace unacceptable chars to acceptable full-width forms
        GetFeasibleFileName: function (name) {
            var ret = name;
            ret = ret.replace(/\\/g, "＼");
            ret = ret.replace(/\|/g, "｜");
            ret = ret.replace(/\//g, "／");
            ret = ret.replace(/\</g, "＜");
            ret = ret.replace(/\>/g, "＞");
            ret = ret.replace(/\"/g, "＂");
            ret = ret.replace(/\*/g, "＊");
            ret = ret.replace(/\?/g, "？");
            ret = ret.replace(/\:/g, "：");
            return ret;
        },
        FirstExistDir: function (lst) {
            for (var i in lst) {
                if (fso.FolderExists(lst[i])) return lst[i];
            }
            return null;
        },
	EnsureDirectoryExist: function (dir) {
		var pos = 0;
		    while (pos != -1) {
		        pos = dir.indexOf("\\", pos + 1);
		        var path = (pos == -1) ? dir : dir.substr(0, pos);
		        try {
		            fso.CreateFolder(path);
		        } catch (e) {  }
		    }
		}
    };
    tps.log = {
        devices: [],
        indent: 0,
        Log: function (level, tag, text) {
            var dt = new Date();
            for (var i in this.devices) {
                this.devices[i].WriteLog(level, tag, dt, this.indent, text);
            }
        },
        Debug: function (text) {
            tps.log.Log("debug", "", text);
        },
        Event: function (text) {
            tps.log.Log("event", "", text);
        },
        Warning: function(text) {
            tps.log.Log("warning", "", text);
        },
        Error: function(text) {
            tps.log.Log("error", "", text);
        },
        Indent: function () {
            this.indent++;
        },
        Unindent: function () {
            this.indent--;
        },
        AddHtmlElementDevice: function(element) {
            var logDevice = {
                WriteLog: function (level, tag, dt, indent, text) {
                    var oLine = document.createElement("span");
                    var timestring = tps.util.FormatDateString(dt, "H:M:S.I");
                    var finaltext = timestring + " " + "".lpad(" ", indent * 2) + text;
                    oLine.appendChild(document.createTextNode(finaltext));
                    oLine.appendChild(document.createElement("br"));
                    oLine.className = "log_" + level;
                    element.appendChild(oLine);
                    element.scrollTop += 1000000;
                }
            };
            tps.log.devices.push(logDevice);
        }
    };
    tps.unittest = {
        ResultOutput: {},
        Expect: function (cond, text) {
            if (cond) {
                this.ResultOutput.OnSuccess(text);
            } else {
                this.ResultOutput.OnFailure(text);
            }
        },
        NewSuite: function (text) {
            this.ResultOutput.NewSuite(text);
        },
        RunSelfTest: function () {
            this.NewSuite("工具函数");
            this.Expect(tps.util.SingleQuote("x") == "'x'", "SingleQuote能正常为单个字符加引号");
            this.Expect(tps.util.SingleQuote("") == "''", "SingleQuote一个空字符串");
            this.Expect(tps.util.RemoveQuote("'x'") == "x", "RemoveQuote可以移除引号");
            this.Expect(tps.util.RemoveQuote("\"'x'\"") == "'x'", "RemoveQuote只移除一层引号");
            this.Expect(tps.util.FormatDateString(new Date(1999, 2, 10, 10, 20, 23, 0), "Y-m-d H:M:S") == "1999-03-10 10:20:23", "FormatDateString格式化普通时间");
            this.Expect(tps.util.arraysEqual([1], [1]), "ArrayEqual, 1 element number array");
            this.Expect(tps.util.arraysEqual([], []), "ArrayEqual, 0 element array");
            this.Expect(tps.util.arraysEqual(["ab", "c", "de"], ["ab", "c", "de"]), "ArrayEqual, 3 elements string array");

            this.NewSuite("文件操作");
            tps.file.WriteTextFileSimple("ice\fire", "t.txt");
            this.Expect(tps.file.GetDir("a.txt") == ".", "GetDir只有文件名，返回当前目录");
            this.Expect(tps.file.GetDir("c:\\a/b\\c.txt/g.t") == "c:\\a/b\\c.txt", "GetDir路径分隔符混合");
            this.Expect(tps.file.ReadTextFileSimple("t.txt") == "ice\fire", "读写文本文件");
            fso.DeleteFile("t.txt");
            this.Expect(tps.file.GetFileName("") == "", "FileName: empty");
            this.Expect(tps.file.GetFileName("a.txt") == "a.txt", "FileName: only name");
            this.Expect(tps.file.GetFileName("c:\\bbbc") == "bbbc", "FileName: normal");

            this.NewSuite("日期");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120122"), "YmdHMS") == "20120122000000", "正确解析yyyymmdd型日期");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120122122334"), "YmdHMS") == "20120122122334", "正确解析yyyymmddHHMMSS型日期");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("now"), "YmdHMS") == tps.util.FormatDateString(new Date(), "YmdHMS"), "正确解析'now'");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("today"), "YmdHMS") == tps.util.FormatDateString(new Date(), "Ymd") + "000000", "正确解析'today'");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222+1m"), "YmdHMS") == "20120322000000", "正确解析相对日期:月份1");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222-121m"), "YmdHMS") == "20020122000000", "正确解析相对日期:月份2");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222+121m"), "YmdHMS") == "20220322000000", "正确解析相对日期:月份3");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222122334+15y"), "YmdHMS") == "20270222122334", "正确解析相对日期:年");

            this.NewSuite("Registry");
            var rootkey = "HKCU\\Software\\tps";
            var key = rootkey + "\\sub1\\sub2";
            tps.reg.CreateKey(key);
            var existed = tps.reg.KeyExisis(key);
            tps.reg.DeleteKey(rootkey);
            this.Expect(existed && !tps.reg.KeyExisis(rootkey), "Reg rootkey creation/deletion");
            tps.reg.SetStringValue(key, "str", "bbb aaa");
            this.Expect(tps.reg.GetStringValue(key, "str") == "bbb aaa", "get/set REG_SZ");
            var multiStringValue = ["12345", "abc  de", "xxxyyy"];
            tps.reg.SetMultiStringValue(key, "multistr", multiStringValue);
            tps.reg.SetMultiStringValue(key, "multistr", multiStringValue); // no exception when overwrite
            var ret = tps.reg.GetMultiStringValue(key, "multistr");
            this.Expect(tps.util.arraysEqual(ret, multiStringValue), "get/set REG_MULTI_SZ");
            tps.reg.SetIntValue(key, "int", 100);
            this.Expect(tps.reg.GetIntValue(key, "int") == 100, "get/set REG_DWORD");
            var subKeys = tps.reg.EnumSubKeys(rootkey + "\\sub1");
            this.Expect(subKeys.length == 1 && subKeys[0].endsWith("sub2"), "EnumKeys");
            var values = tps.reg.EnumValues(key);
            this.Expect(values.length == 3 && values[2].valstr == "0x64" && values[1].type == "REG_MULTI_SZ" && values[0].name == "str", "EnumValues");
            tps.reg.DeleteKey(rootkey);
        }
    };
}());
