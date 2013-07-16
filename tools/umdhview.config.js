// Tell umdhview, whether given source file is our source
// This can help umdhview to identify the culprit frame,
// because the first frame which belongs to our source is most likely culprit
// param:   [in] sourcefile : source file full path
// return:  true:  indicates sourcefile is ours
//          false: otherwise
IsMySource = function(sourcefile) {
            return sourcefile.search(/enhancedupp/i) != -1;
        };

// Tell umdhview, whether given culprit should be ignored
// param:   [in] frame : culprit frame, an object with following string properties
//                       {module, func, offset, source, line}
// return:  true:  indicates this culprit should be ignored
//          false: otherwise
// note:    All culprits shall not be ignored unless it's part of caching or it's by design
//          Any modification of this function should be approved by euppeng@microsoft.com
IsCulpritIgnored = function(frame) {
	return false;
};
