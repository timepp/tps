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

(function () {

    var ForReading = 1, ForWriting = 2;
    var g_wmi = null;

    function WMI(path) {
        if (!g_wmi) {
            g_wmi = new Object;
        }
        if (!g_wmi[path]) {
            g_wmi[path] = GetObject("winmgmts:root/" + path);
        }
        return g_wmi[path];
    }
    function REG() {
        return WMI("default").Get("StdRegProv");
    }
    function InvokeCommonRegTask(cmd, root, key, valname, val) {
        var func = REG().Methods_.Item(cmd);
        var param = func.InParameters.SpawnInstance_();
        param.hDefKey = root;
        param.sSubKeyName = key;
        param.sValueName = valname;
        if (val != null) {
            if (typeof (val) == "string") param.sValue = val;
            else param.uValue = val;
        }
        return REG().ExecMethod_(func.Name, param);
    }

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
        IndexOf: function (arr, x) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] == x) return i;
            }
            return -1;
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
            var D2 = function (n) { if (n < 10) return "0" + n.toString(); return n.toString(); };
            for (var i = 0; i < fmt.length; i++) {
                var c = fmt.charAt(i);
                if (c == "Y") ret += dt.getFullYear();
                else if (c == "m") ret += D2(dt.getMonth() + 1);
                else if (c == "d") ret += D2(dt.getDate());
                else if (c == "H") ret += D2(dt.getHours());
                else if (c == "M") ret += D2(dt.getMinutes());
                else if (c == "S") ret += D2(dt.getSeconds());
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
                        return new Date(m[1], m[2] - 1, m[3], m[5], m[6], m[7], 0);
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
        }
    };

    tps.sys = {
        HasFullPrivilege: function () {
            try {
                var value = shell.RegRead("HKEY_USERS\\S-1-5-19\\");
            } catch (e) {
                return false;
            }
            return true;
        },
        IsAdmin: function () {
            var oNet = new ActiveXObject("WScript.Network");
            var oGroup = GetObject("WinNT://./Administrators");
            var e = new Enumerator(oGroup.Members());
            for (; !e.atEnd() ; e.moveNext()) {
                if (e.item().Name == oNet.UserName) return true;
            }
            return false;
        },
        GetScriptPath: function () {
            try {
                return WScript.ScriptFullName;
            } catch (e) {
                return document.location.pathname;
            }
        },
        GetScriptDir: function () {
            return tps.file.GetDir(tps.sys.GetScriptPath());
        },
        GetSystemEnv: function (vname) {
            var items = WMI("cimv2").ExecQuery("Select * from Win32_Environment Where Name = '$V'".replace("$V", vname));
            if (!items || items.Count == 0) return null;
            var item = new Enumerator(items).item();
            return item.VariableValue;
        },
        SetSystemEnv: function (vname, val, login_name) {
            if (login_name == null) login_name = "<SYSTEM>";

            var item = WMI("cimv2").Get("Win32_Environment").SpawnInstance_();
            item.Name = vname;
            item.Username = login_name;
            item.VariableValue = val;
            item.Put_();

            // Broadcast a "environment change" notification, 
            // so that processes created by explorer.exe can see the new variable before the next login.
            tps.sys.RunCommandAndGetResult(tps.sys.GetScriptDir() + "\\tpkit.exe --action=BroadcastEnvironmentChange");
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
        RunCommandAndGetResult: function (cmdline, of, ef) {
            var outfile = of ? of : shell.ExpandEnvironmentStrings("%temp%") + "\\" + fso.GetTempName();
            var errfile = ef ? ef : shell.ExpandEnvironmentStrings("%temp%") + "\\" + fso.GetTempName();

            try {
                fso.DeleteFile(outfile);
                fso.DeleteFile(errfile);
            } catch (e) { }

            cmdline = "cmd.exe /C " + cmdline + ' > "OUT" 2> "ERR"'.replace("OUT", outfile).replace("ERR", errfile);
            Log("RUN:[" + cmdline + "]");
            shell.Run(cmdline, 0, true);
            var ret = {
                output: tps.file.ReadTextFileSimple(outfile),
                errors: tps.file.ReadTextFileSimple(errfile)
            };

            try {
                if (!of) fso.DeleteFile(outfile);
                if (!ef) fso.DeleteFile(errfile);
            } catch (e) {}

            return ret;
        }
    };

    tps.reg = {
        GetStringValue: function (root, key, val) {
            InvokeCommonRegTask("GetStringValue", root, key, val).sValue;
        },
        SetStringValue: function (root, key, valname, val) {
            return InvokeCommonRegTask("SetStringValue", root, key, valname, val);
        },
        GetIntValue: function (root, key, val) {
            return InvokeCommonRegTask("GetDWORDValue", root, key, val).uValue;
        },
        SetIntValue: function (root, key, valname, val) {
            return InvokeCommonRegTask("SetDWORDValue", root, key, valname, val);
        },
        StringValueExists: function (root, key, val) {
            var s = GetStringValue(root, key, val);
            return s != undefined && s != null;
        },
        IntValueExists: function (root, key, val) {
            var s = GetIntValue(root, key, val);
            return s != undefined && s != null;
        },
        OpenRegEdit: function (path) {
            tps.reg.SetStringValue(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Applets\\Regedit", "LastKey", path);
            shell.Run("regedit.exe");
        }
    };

    tps.ui = {
        ResizeWindow: function (cx, cy, center) {
            window.resizeTo(cx, cy);
            if (center) {
                var items = WMI("cimv2").ExecQuery("Select * From Win32_DesktopMonitor");
                var item = new Enumerator(items).item();
                var w = item.ScreenWidth;
                var h = item.ScreenHeight;
                window.moveTo((w - cx) / 2, (h - cy) / 2);
            }
        },
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
            //	stream.Type = 1;
            stream.Open();
            stream.LoadFromFile(filename);
            stream.Position = pos;
            var str = "";
            //	str = stream.Read(len);
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

            this.NewSuite("文件操作");
            tps.file.WriteTextFileSimple("ice\fire", "t.txt");
            this.Expect(tps.file.GetDir("a.txt") == ".", "GetDir只有文件名，返回当前目录");
            this.Expect(tps.file.GetDir("c:\\a/b\\c.txt/g.t") == "c:\\a/b\\c.txt", "GetDir路径分隔符混合");
            this.Expect(tps.file.ReadTextFileSimple("t.txt") == "ice\fire", "读写文本文件");
            fso.DeleteFile("t.txt");

            this.NewSuite("日期");
            debugger;
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120122"), "YmdHMS") == "20120122000000", "正确解析yyyymmdd型日期");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120122122334"), "YmdHMS") == "20120122122334", "正确解析yyyymmddHHMMSS型日期");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("now"), "YmdHMS") == tps.util.FormatDateString(new Date(), "YmdHMS"), "正确解析'now'");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("today"), "YmdHMS") == tps.util.FormatDateString(new Date(), "Ymd") + "000000", "正确解析'today'");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222+1m"), "YmdHMS") == "20120322000000", "正确解析相对日期:月份1");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222-121m"), "YmdHMS") == "20020122000000", "正确解析相对日期:月份2");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222+121m"), "YmdHMS") == "20220322000000", "正确解析相对日期:月份3");
            this.Expect(tps.util.FormatDateString(tps.util.ParseDateString("20120222122334+15y"), "YmdHMS") == "20270222122334", "正确解析相对日期:年");
        }
    };
}());
