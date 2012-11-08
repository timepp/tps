if (!this.tps) {
	this.tps = {};
}

// wellknown windows object
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");
var shellapp = new ActiveXObject("Shell.Application");
var env = shell.Environment("Process");

var HKEY_LOCAL_MACHINE = 0x80000002;
var HKEY_CURRENT_USER  = 0x80000001;
var HKEY_CLASSES_ROOT  = 0x80000000;

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
				if (!pn in obj) obj[pn] = (pn == arg ? new Object : new Array);
				obj = obj[pn];
			}
			return obj;
		},
		FormatDateString: function (dt, fmt) {
			var ret = "";
			var D2 = function (n) { if (n < 10) return "0" + n.toString(); return n.toString(); };
			for (var i = 0; i < fmt.length; i++) {
				var c = fmt.charAt(i);
				if (c == "Y") ret += dt.getFullYear();
				else if (c == "m") ret += D2(dt.getMonth());
				else if (c == "d") ret += D2(dt.getDate());
				else if (c == "H") ret += D2(dt.getHours());
				else if (c == "M") ret += D2(dt.getMinutes());
				else if (c == "S") ret += D2(dt.getSeconds());
				else ret += c;
			}
			return ret;
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

			cmdline += ' > "OUT" 2> "ERR"'.replace("OUT", outfile).replace("ERR", errfile);
			Log("RUN:[" + cmdline + "]");
			shell.Run(cmdline, 0, true);
			var ret = {
				output: GetTextFileContent(outfile),
				errors: GetTextFileContent(errfile)
			};

			if (!of) fso.DeleteFile(outfile);
			if (!ef) fso.DeleteFile(errfile);

			return ret;
		}
	};



	//function ElevatePrivilege(cmdline) {
	//	if (!HasFullPrivilege()) {
	//		var oNet = new ActiveXObject("WScript.Network");
	//		var ret = shellapp.ShellExecute("mshta.exe", cmdline + "--uac --user=" + oNet.UserName, "", "runas", 1);
	//		window.close();
	//		return true;
	//	}
	//	return false;
	//}



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
			oSpan.innerHTML = "<button onclick=\"document.body.removeChild(document.getElementById('div_err'))\">关闭</button>";
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
		CreateShortcut: function (lnkName, targetPath, dir, argument, desc) {
			var oLnk = shell.CreateShortcut(lnkName);
			oLnk.TargetPath = targetPath;
			oLnk.WorkingDirectory = dir;
			oLnk.Description = desc;
			oLnk.Arguments = argument;
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
		Glob: function (dir, filter) {
			var ret = { dirs: [], files: [] };
			var d = fso.GetFolder(dir);
			for (var fc = new Enumerator(d.files) ; !fc.atEnd() ; fc.moveNext()) {
				var f = fc.item();
				if (!filter || filter.test(f.Name)) {
					ret.files.push(f.Name);
				}
			}
			for (var fc = new Enumerator(d.SubFolders) ; !fc.atEnd() ; fc.moveNext()) {
				var f = fc.item();
				if (!filter || filter.test(f.Name)) {
					ret.dirs.push(f.Name);
				}
			}
			return ret;
		},
		GetDir: function (path) {
			return path.replace(/^(.*)[\\/][^\\/]+$/, "$1");
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
			this.Expect(tps.util.FormatDateString(new Date(1999, 2, 10, 10, 20, 23, 0), "Y-m-d H:M:S") == "1999-02-10 10:20:23", "FormatDateString格式化普通时间");
		}
	};
}());
