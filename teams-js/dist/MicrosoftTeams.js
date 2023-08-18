(function webpackUniversalModuleDefinition(root, factory) {
	if(typeof exports === 'object' && typeof module === 'object')
		module.exports = factory();
	else if(typeof define === 'function' && define.amd)
		define("microsoftTeams", [], factory);
	else if(typeof exports === 'object')
		exports["microsoftTeams"] = factory();
	else
		root["microsoftTeams"] = factory();
})(typeof self !== 'undefined' ? self : this, () => {
return /******/ (() => { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ 302:
/***/ ((module, exports, __webpack_require__) => {

/* eslint-env browser */

/**
 * This is the web browser implementation of `debug()`.
 */

exports.formatArgs = formatArgs;
exports.save = save;
exports.load = load;
exports.useColors = useColors;
exports.storage = localstorage();
exports.destroy = (() => {
	let warned = false;

	return () => {
		if (!warned) {
			warned = true;
			console.warn('Instance method `debug.destroy()` is deprecated and no longer does anything. It will be removed in the next major version of `debug`.');
		}
	};
})();

/**
 * Colors.
 */

exports.colors = [
	'#0000CC',
	'#0000FF',
	'#0033CC',
	'#0033FF',
	'#0066CC',
	'#0066FF',
	'#0099CC',
	'#0099FF',
	'#00CC00',
	'#00CC33',
	'#00CC66',
	'#00CC99',
	'#00CCCC',
	'#00CCFF',
	'#3300CC',
	'#3300FF',
	'#3333CC',
	'#3333FF',
	'#3366CC',
	'#3366FF',
	'#3399CC',
	'#3399FF',
	'#33CC00',
	'#33CC33',
	'#33CC66',
	'#33CC99',
	'#33CCCC',
	'#33CCFF',
	'#6600CC',
	'#6600FF',
	'#6633CC',
	'#6633FF',
	'#66CC00',
	'#66CC33',
	'#9900CC',
	'#9900FF',
	'#9933CC',
	'#9933FF',
	'#99CC00',
	'#99CC33',
	'#CC0000',
	'#CC0033',
	'#CC0066',
	'#CC0099',
	'#CC00CC',
	'#CC00FF',
	'#CC3300',
	'#CC3333',
	'#CC3366',
	'#CC3399',
	'#CC33CC',
	'#CC33FF',
	'#CC6600',
	'#CC6633',
	'#CC9900',
	'#CC9933',
	'#CCCC00',
	'#CCCC33',
	'#FF0000',
	'#FF0033',
	'#FF0066',
	'#FF0099',
	'#FF00CC',
	'#FF00FF',
	'#FF3300',
	'#FF3333',
	'#FF3366',
	'#FF3399',
	'#FF33CC',
	'#FF33FF',
	'#FF6600',
	'#FF6633',
	'#FF9900',
	'#FF9933',
	'#FFCC00',
	'#FFCC33'
];

/**
 * Currently only WebKit-based Web Inspectors, Firefox >= v31,
 * and the Firebug extension (any Firefox version) are known
 * to support "%c" CSS customizations.
 *
 * TODO: add a `localStorage` variable to explicitly enable/disable colors
 */

// eslint-disable-next-line complexity
function useColors() {
	// NB: In an Electron preload script, document will be defined but not fully
	// initialized. Since we know we're in Chrome, we'll just detect this case
	// explicitly
	if (typeof window !== 'undefined' && window.process && (window.process.type === 'renderer' || window.process.__nwjs)) {
		return true;
	}

	// Internet Explorer and Edge do not support colors.
	if (typeof navigator !== 'undefined' && navigator.userAgent && navigator.userAgent.toLowerCase().match(/(edge|trident)\/(\d+)/)) {
		return false;
	}

	// Is webkit? http://stackoverflow.com/a/16459606/376773
	// document is undefined in react-native: https://github.com/facebook/react-native/pull/1632
	return (typeof document !== 'undefined' && document.documentElement && document.documentElement.style && document.documentElement.style.WebkitAppearance) ||
		// Is firebug? http://stackoverflow.com/a/398120/376773
		(typeof window !== 'undefined' && window.console && (window.console.firebug || (window.console.exception && window.console.table))) ||
		// Is firefox >= v31?
		// https://developer.mozilla.org/en-US/docs/Tools/Web_Console#Styling_messages
		(typeof navigator !== 'undefined' && navigator.userAgent && navigator.userAgent.toLowerCase().match(/firefox\/(\d+)/) && parseInt(RegExp.$1, 10) >= 31) ||
		// Double check webkit in userAgent just in case we are in a worker
		(typeof navigator !== 'undefined' && navigator.userAgent && navigator.userAgent.toLowerCase().match(/applewebkit\/(\d+)/));
}

/**
 * Colorize log arguments if enabled.
 *
 * @api public
 */

function formatArgs(args) {
	args[0] = (this.useColors ? '%c' : '') +
		this.namespace +
		(this.useColors ? ' %c' : ' ') +
		args[0] +
		(this.useColors ? '%c ' : ' ') +
		'+' + module.exports.humanize(this.diff);

	if (!this.useColors) {
		return;
	}

	const c = 'color: ' + this.color;
	args.splice(1, 0, c, 'color: inherit');

	// The final "%c" is somewhat tricky, because there could be other
	// arguments passed either before or after the %c, so we need to
	// figure out the correct index to insert the CSS into
	let index = 0;
	let lastC = 0;
	args[0].replace(/%[a-zA-Z%]/g, match => {
		if (match === '%%') {
			return;
		}
		index++;
		if (match === '%c') {
			// We only are interested in the *last* %c
			// (the user may have provided their own)
			lastC = index;
		}
	});

	args.splice(lastC, 0, c);
}

/**
 * Invokes `console.debug()` when available.
 * No-op when `console.debug` is not a "function".
 * If `console.debug` is not available, falls back
 * to `console.log`.
 *
 * @api public
 */
exports.log = console.debug || console.log || (() => {});

/**
 * Save `namespaces`.
 *
 * @param {String} namespaces
 * @api private
 */
function save(namespaces) {
	try {
		if (namespaces) {
			exports.storage.setItem('debug', namespaces);
		} else {
			exports.storage.removeItem('debug');
		}
	} catch (error) {
		// Swallow
		// XXX (@Qix-) should we be logging these?
	}
}

/**
 * Load `namespaces`.
 *
 * @return {String} returns the previously persisted debug modes
 * @api private
 */
function load() {
	let r;
	try {
		r = exports.storage.getItem('debug');
	} catch (error) {
		// Swallow
		// XXX (@Qix-) should we be logging these?
	}

	// If debug isn't set in LS, and we're in Electron, try to load $DEBUG
	if (!r && typeof process !== 'undefined' && 'env' in process) {
		r = process.env.DEBUG;
	}

	return r;
}

/**
 * Localstorage attempts to return the localstorage.
 *
 * This is necessary because safari throws
 * when a user disables cookies/localstorage
 * and you attempt to access it.
 *
 * @return {LocalStorage}
 * @api private
 */

function localstorage() {
	try {
		// TVMLKit (Apple TV JS Runtime) does not have a window object, just localStorage in the global context
		// The Browser also has localStorage in the global context.
		return localStorage;
	} catch (error) {
		// Swallow
		// XXX (@Qix-) should we be logging these?
	}
}

module.exports = __webpack_require__(65)(exports);

const {formatters} = module.exports;

/**
 * Map %j to `JSON.stringify()`, since no Web Inspectors do that by default.
 */

formatters.j = function (v) {
	try {
		return JSON.stringify(v);
	} catch (error) {
		return '[UnexpectedJSONParseError]: ' + error.message;
	}
};


/***/ }),

/***/ 65:
/***/ ((module, __unused_webpack_exports, __webpack_require__) => {


/**
 * This is the common logic for both the Node.js and web browser
 * implementations of `debug()`.
 */

function setup(env) {
	createDebug.debug = createDebug;
	createDebug.default = createDebug;
	createDebug.coerce = coerce;
	createDebug.disable = disable;
	createDebug.enable = enable;
	createDebug.enabled = enabled;
	createDebug.humanize = __webpack_require__(247);
	createDebug.destroy = destroy;

	Object.keys(env).forEach(key => {
		createDebug[key] = env[key];
	});

	/**
	* The currently active debug mode names, and names to skip.
	*/

	createDebug.names = [];
	createDebug.skips = [];

	/**
	* Map of special "%n" handling functions, for the debug "format" argument.
	*
	* Valid key names are a single, lower or upper-case letter, i.e. "n" and "N".
	*/
	createDebug.formatters = {};

	/**
	* Selects a color for a debug namespace
	* @param {String} namespace The namespace string for the debug instance to be colored
	* @return {Number|String} An ANSI color code for the given namespace
	* @api private
	*/
	function selectColor(namespace) {
		let hash = 0;

		for (let i = 0; i < namespace.length; i++) {
			hash = ((hash << 5) - hash) + namespace.charCodeAt(i);
			hash |= 0; // Convert to 32bit integer
		}

		return createDebug.colors[Math.abs(hash) % createDebug.colors.length];
	}
	createDebug.selectColor = selectColor;

	/**
	* Create a debugger with the given `namespace`.
	*
	* @param {String} namespace
	* @return {Function}
	* @api public
	*/
	function createDebug(namespace) {
		let prevTime;
		let enableOverride = null;
		let namespacesCache;
		let enabledCache;

		function debug(...args) {
			// Disabled?
			if (!debug.enabled) {
				return;
			}

			const self = debug;

			// Set `diff` timestamp
			const curr = Number(new Date());
			const ms = curr - (prevTime || curr);
			self.diff = ms;
			self.prev = prevTime;
			self.curr = curr;
			prevTime = curr;

			args[0] = createDebug.coerce(args[0]);

			if (typeof args[0] !== 'string') {
				// Anything else let's inspect with %O
				args.unshift('%O');
			}

			// Apply any `formatters` transformations
			let index = 0;
			args[0] = args[0].replace(/%([a-zA-Z%])/g, (match, format) => {
				// If we encounter an escaped % then don't increase the array index
				if (match === '%%') {
					return '%';
				}
				index++;
				const formatter = createDebug.formatters[format];
				if (typeof formatter === 'function') {
					const val = args[index];
					match = formatter.call(self, val);

					// Now we need to remove `args[index]` since it's inlined in the `format`
					args.splice(index, 1);
					index--;
				}
				return match;
			});

			// Apply env-specific formatting (colors, etc.)
			createDebug.formatArgs.call(self, args);

			const logFn = self.log || createDebug.log;
			logFn.apply(self, args);
		}

		debug.namespace = namespace;
		debug.useColors = createDebug.useColors();
		debug.color = createDebug.selectColor(namespace);
		debug.extend = extend;
		debug.destroy = createDebug.destroy; // XXX Temporary. Will be removed in the next major release.

		Object.defineProperty(debug, 'enabled', {
			enumerable: true,
			configurable: false,
			get: () => {
				if (enableOverride !== null) {
					return enableOverride;
				}
				if (namespacesCache !== createDebug.namespaces) {
					namespacesCache = createDebug.namespaces;
					enabledCache = createDebug.enabled(namespace);
				}

				return enabledCache;
			},
			set: v => {
				enableOverride = v;
			}
		});

		// Env-specific initialization logic for debug instances
		if (typeof createDebug.init === 'function') {
			createDebug.init(debug);
		}

		return debug;
	}

	function extend(namespace, delimiter) {
		const newDebug = createDebug(this.namespace + (typeof delimiter === 'undefined' ? ':' : delimiter) + namespace);
		newDebug.log = this.log;
		return newDebug;
	}

	/**
	* Enables a debug mode by namespaces. This can include modes
	* separated by a colon and wildcards.
	*
	* @param {String} namespaces
	* @api public
	*/
	function enable(namespaces) {
		createDebug.save(namespaces);
		createDebug.namespaces = namespaces;

		createDebug.names = [];
		createDebug.skips = [];

		let i;
		const split = (typeof namespaces === 'string' ? namespaces : '').split(/[\s,]+/);
		const len = split.length;

		for (i = 0; i < len; i++) {
			if (!split[i]) {
				// ignore empty strings
				continue;
			}

			namespaces = split[i].replace(/\*/g, '.*?');

			if (namespaces[0] === '-') {
				createDebug.skips.push(new RegExp('^' + namespaces.slice(1) + '$'));
			} else {
				createDebug.names.push(new RegExp('^' + namespaces + '$'));
			}
		}
	}

	/**
	* Disable debug output.
	*
	* @return {String} namespaces
	* @api public
	*/
	function disable() {
		const namespaces = [
			...createDebug.names.map(toNamespace),
			...createDebug.skips.map(toNamespace).map(namespace => '-' + namespace)
		].join(',');
		createDebug.enable('');
		return namespaces;
	}

	/**
	* Returns true if the given mode name is enabled, false otherwise.
	*
	* @param {String} name
	* @return {Boolean}
	* @api public
	*/
	function enabled(name) {
		if (name[name.length - 1] === '*') {
			return true;
		}

		let i;
		let len;

		for (i = 0, len = createDebug.skips.length; i < len; i++) {
			if (createDebug.skips[i].test(name)) {
				return false;
			}
		}

		for (i = 0, len = createDebug.names.length; i < len; i++) {
			if (createDebug.names[i].test(name)) {
				return true;
			}
		}

		return false;
	}

	/**
	* Convert regexp to namespace
	*
	* @param {RegExp} regxep
	* @return {String} namespace
	* @api private
	*/
	function toNamespace(regexp) {
		return regexp.toString()
			.substring(2, regexp.toString().length - 2)
			.replace(/\.\*\?$/, '*');
	}

	/**
	* Coerce `val`.
	*
	* @param {Mixed} val
	* @return {Mixed}
	* @api private
	*/
	function coerce(val) {
		if (val instanceof Error) {
			return val.stack || val.message;
		}
		return val;
	}

	/**
	* XXX DO NOT USE. This is a temporary stub function.
	* XXX It WILL be removed in the next major release.
	*/
	function destroy() {
		console.warn('Instance method `debug.destroy()` is deprecated and no longer does anything. It will be removed in the next major version of `debug`.');
	}

	createDebug.enable(createDebug.load());

	return createDebug;
}

module.exports = setup;


/***/ }),

/***/ 247:
/***/ ((module) => {

/**
 * Helpers.
 */

var s = 1000;
var m = s * 60;
var h = m * 60;
var d = h * 24;
var w = d * 7;
var y = d * 365.25;

/**
 * Parse or format the given `val`.
 *
 * Options:
 *
 *  - `long` verbose formatting [false]
 *
 * @param {String|Number} val
 * @param {Object} [options]
 * @throws {Error} throw an error if val is not a non-empty string or a number
 * @return {String|Number}
 * @api public
 */

module.exports = function(val, options) {
  options = options || {};
  var type = typeof val;
  if (type === 'string' && val.length > 0) {
    return parse(val);
  } else if (type === 'number' && isFinite(val)) {
    return options.long ? fmtLong(val) : fmtShort(val);
  }
  throw new Error(
    'val is not a non-empty string or a valid number. val=' +
      JSON.stringify(val)
  );
};

/**
 * Parse the given `str` and return milliseconds.
 *
 * @param {String} str
 * @return {Number}
 * @api private
 */

function parse(str) {
  str = String(str);
  if (str.length > 100) {
    return;
  }
  var match = /^(-?(?:\d+)?\.?\d+) *(milliseconds?|msecs?|ms|seconds?|secs?|s|minutes?|mins?|m|hours?|hrs?|h|days?|d|weeks?|w|years?|yrs?|y)?$/i.exec(
    str
  );
  if (!match) {
    return;
  }
  var n = parseFloat(match[1]);
  var type = (match[2] || 'ms').toLowerCase();
  switch (type) {
    case 'years':
    case 'year':
    case 'yrs':
    case 'yr':
    case 'y':
      return n * y;
    case 'weeks':
    case 'week':
    case 'w':
      return n * w;
    case 'days':
    case 'day':
    case 'd':
      return n * d;
    case 'hours':
    case 'hour':
    case 'hrs':
    case 'hr':
    case 'h':
      return n * h;
    case 'minutes':
    case 'minute':
    case 'mins':
    case 'min':
    case 'm':
      return n * m;
    case 'seconds':
    case 'second':
    case 'secs':
    case 'sec':
    case 's':
      return n * s;
    case 'milliseconds':
    case 'millisecond':
    case 'msecs':
    case 'msec':
    case 'ms':
      return n;
    default:
      return undefined;
  }
}

/**
 * Short format for `ms`.
 *
 * @param {Number} ms
 * @return {String}
 * @api private
 */

function fmtShort(ms) {
  var msAbs = Math.abs(ms);
  if (msAbs >= d) {
    return Math.round(ms / d) + 'd';
  }
  if (msAbs >= h) {
    return Math.round(ms / h) + 'h';
  }
  if (msAbs >= m) {
    return Math.round(ms / m) + 'm';
  }
  if (msAbs >= s) {
    return Math.round(ms / s) + 's';
  }
  return ms + 'ms';
}

/**
 * Long format for `ms`.
 *
 * @param {Number} ms
 * @return {String}
 * @api private
 */

function fmtLong(ms) {
  var msAbs = Math.abs(ms);
  if (msAbs >= d) {
    return plural(ms, msAbs, d, 'day');
  }
  if (msAbs >= h) {
    return plural(ms, msAbs, h, 'hour');
  }
  if (msAbs >= m) {
    return plural(ms, msAbs, m, 'minute');
  }
  if (msAbs >= s) {
    return plural(ms, msAbs, s, 'second');
  }
  return ms + ' ms';
}

/**
 * Pluralization helper.
 */

function plural(ms, msAbs, n, name) {
  var isPlural = msAbs >= n * 1.5;
  return Math.round(ms / n) + ' ' + name + (isPlural ? 's' : '');
}


/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be in strict mode.
(() => {
"use strict";
// ESM COMPAT FLAG
__webpack_require__.r(__webpack_exports__);

// EXPORTS
__webpack_require__.d(__webpack_exports__, {
  "ActionObjectType": () => (/* reexport */ ActionObjectType),
  "ChannelType": () => (/* reexport */ ChannelType),
  "ChildAppWindow": () => (/* reexport */ ChildAppWindow),
  "DialogDimension": () => (/* reexport */ DialogDimension),
  "ErrorCode": () => (/* reexport */ ErrorCode),
  "FileOpenPreference": () => (/* reexport */ FileOpenPreference),
  "FrameContexts": () => (/* reexport */ FrameContexts),
  "HostClientType": () => (/* reexport */ HostClientType),
  "HostName": () => (/* reexport */ HostName),
  "LiveShareHost": () => (/* reexport */ LiveShareHost),
  "NotificationTypes": () => (/* reexport */ NotificationTypes),
  "ParentAppWindow": () => (/* reexport */ ParentAppWindow),
  "SecondaryM365ContentIdName": () => (/* reexport */ SecondaryM365ContentIdName),
  "TaskModuleDimension": () => (/* reexport */ TaskModuleDimension),
  "TeamType": () => (/* reexport */ TeamType),
  "UserSettingTypes": () => (/* reexport */ UserSettingTypes),
  "UserTeamRole": () => (/* reexport */ UserTeamRole),
  "ViewerActionTypes": () => (/* reexport */ ViewerActionTypes),
  "app": () => (/* reexport */ app),
  "appEntity": () => (/* reexport */ appEntity),
  "appInitialization": () => (/* reexport */ appInitialization),
  "appInstallDialog": () => (/* reexport */ appInstallDialog),
  "appNotification": () => (/* reexport */ appNotification),
  "authentication": () => (/* reexport */ authentication),
  "barCode": () => (/* reexport */ barCode),
  "calendar": () => (/* reexport */ calendar),
  "call": () => (/* reexport */ call),
  "chat": () => (/* reexport */ chat),
  "conversations": () => (/* reexport */ conversations),
  "dialog": () => (/* reexport */ dialog),
  "enablePrintCapability": () => (/* reexport */ enablePrintCapability),
  "executeDeepLink": () => (/* reexport */ executeDeepLink),
  "files": () => (/* reexport */ files),
  "geoLocation": () => (/* reexport */ geoLocation),
  "getAdaptiveCardSchemaVersion": () => (/* reexport */ getAdaptiveCardSchemaVersion),
  "getContext": () => (/* reexport */ getContext),
  "getMruTabInstances": () => (/* reexport */ getMruTabInstances),
  "getTabInstances": () => (/* reexport */ getTabInstances),
  "initialize": () => (/* reexport */ initialize),
  "initializeWithFrameContext": () => (/* reexport */ initializeWithFrameContext),
  "liveShare": () => (/* reexport */ liveShare),
  "location": () => (/* reexport */ location_location),
  "logs": () => (/* reexport */ logs),
  "mail": () => (/* reexport */ mail),
  "marketplace": () => (/* reexport */ marketplace),
  "media": () => (/* reexport */ media),
  "meeting": () => (/* reexport */ meeting),
  "meetingRoom": () => (/* reexport */ meetingRoom),
  "menus": () => (/* reexport */ menus),
  "monetization": () => (/* reexport */ monetization),
  "navigateBack": () => (/* reexport */ navigateBack),
  "navigateCrossDomain": () => (/* reexport */ navigateCrossDomain),
  "navigateToTab": () => (/* reexport */ navigateToTab),
  "notifications": () => (/* reexport */ notifications),
  "openFilePreview": () => (/* reexport */ openFilePreview),
  "pages": () => (/* reexport */ pages),
  "people": () => (/* reexport */ people),
  "print": () => (/* reexport */ print),
  "profile": () => (/* reexport */ profile),
  "registerAppButtonClickHandler": () => (/* reexport */ registerAppButtonClickHandler),
  "registerAppButtonHoverEnterHandler": () => (/* reexport */ registerAppButtonHoverEnterHandler),
  "registerAppButtonHoverLeaveHandler": () => (/* reexport */ registerAppButtonHoverLeaveHandler),
  "registerBackButtonHandler": () => (/* reexport */ registerBackButtonHandler),
  "registerBeforeUnloadHandler": () => (/* reexport */ registerBeforeUnloadHandler),
  "registerChangeSettingsHandler": () => (/* reexport */ registerChangeSettingsHandler),
  "registerCustomHandler": () => (/* reexport */ registerCustomHandler),
  "registerFocusEnterHandler": () => (/* reexport */ registerFocusEnterHandler),
  "registerFullScreenHandler": () => (/* reexport */ registerFullScreenHandler),
  "registerOnLoadHandler": () => (/* reexport */ registerOnLoadHandler),
  "registerOnThemeChangeHandler": () => (/* reexport */ registerOnThemeChangeHandler),
  "registerUserSettingsChangeHandler": () => (/* reexport */ registerUserSettingsChangeHandler),
  "remoteCamera": () => (/* reexport */ remoteCamera),
  "returnFocus": () => (/* reexport */ returnFocus),
  "search": () => (/* reexport */ search),
  "secondaryBrowser": () => (/* reexport */ secondaryBrowser),
  "sendCustomEvent": () => (/* reexport */ sendCustomEvent),
  "sendCustomMessage": () => (/* reexport */ sendCustomMessage),
  "setFrameContext": () => (/* reexport */ setFrameContext),
  "settings": () => (/* reexport */ settings),
  "shareDeepLink": () => (/* reexport */ shareDeepLink),
  "sharing": () => (/* reexport */ sharing),
  "stageView": () => (/* reexport */ stageView),
  "tasks": () => (/* reexport */ tasks),
  "teams": () => (/* reexport */ teams),
  "teamsCore": () => (/* reexport */ teamsCore),
  "uploadCustomApp": () => (/* reexport */ uploadCustomApp),
  "version": () => (/* reexport */ version),
  "video": () => (/* reexport */ video),
  "videoEx": () => (/* reexport */ videoEx),
  "webStorage": () => (/* reexport */ webStorage)
});

;// CONCATENATED MODULE: ./src/internal/constants.ts
/**
 * @hidden
 * The client version when all SDK APIs started to check platform compatibility for the APIs was 1.6.0.
 * Modified to 2.0.1 which is hightest till now so that if any client doesn't pass version in initialize function, it will be set to highest.
 * Mobile clients are passing versions, hence will be applicable to web and desktop clients only.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var defaultSDKVersionForCompatCheck = '2.0.1';
/**
 * @hidden
 * This is the client version when selectMedia API - VideoAndImage is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var videoAndImageMediaAPISupportVersion = '2.0.2';
/**
 * @hidden
 * This is the client version when selectMedia API - Video with non-full screen mode is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var nonFullScreenVideoModeAPISupportVersion = '2.0.3';
/**
 * @hidden
 * This is the client version when selectMedia API - ImageOutputFormats is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var imageOutputFormatsAPISupportVersion = '2.0.4';
/**
 * @hidden
 * Minimum required client supported version for {@link getUserJoinedTeams} to be supported on {@link HostClientType.android}
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var getUserJoinedTeamsSupportedAndroidClientVersion = '2.0.1';
/**
 * @hidden
 * This is the client version when location APIs (getLocation and showLocation) are supported.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var locationAPIsRequiredVersion = '1.9.0';
/**
 * @hidden
 * This is the client version when permisisons are supported
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var permissionsAPIsRequiredVersion = '2.0.1';
/**
 * @hidden
 * This is the client version when people picker API is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var peoplePickerRequiredVersion = '2.0.0';
/**
 * @hidden
 * This is the client version when captureImage API is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var captureImageMobileSupportVersion = '1.7.0';
/**
 * @hidden
 * This is the client version when media APIs are supported on all three platforms ios, android and web.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var mediaAPISupportVersion = '1.8.0';
/**
 * @hidden
 * This is the client version when getMedia API is supported via Callbacks on all three platforms ios, android and web.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var getMediaCallbackSupportVersion = '2.0.0';
/**
 * @hidden
 * This is the client version when scanBarCode API is supported on mobile.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var scanBarCodeAPIMobileSupportVersion = '1.9.0';
/**
 * @hidden
 * List of supported Host origins
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var validOrigins = [
    'teams.microsoft.com',
    'teams.microsoft.us',
    'gov.teams.microsoft.us',
    'dod.teams.microsoft.us',
    'int.teams.microsoft.com',
    'teams.live.com',
    'devspaces.skype.com',
    'ssauth.skype.com',
    'local.teams.live.com',
    'local.teams.live.com:8080',
    'local.teams.office.com',
    'local.teams.office.com:8080',
    'outlook.office.com',
    'outlook-sdf.office.com',
    'outlook.office365.com',
    'outlook-sdf.office365.com',
    'outlook.live.com',
    'outlook-sdf.live.com',
    '*.teams.microsoft.com',
    '*.www.office.com',
    'www.office.com',
    'word.office.com',
    'excel.office.com',
    'powerpoint.office.com',
    'www.officeppe.com',
    '*.www.microsoft365.com',
    'www.microsoft365.com',
];
/**
 * @hidden
 * USer specified message origins should satisfy this test
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var userOriginUrlValidationRegExp = /^https:\/\//;
/**
 * @hidden
 * The protocol used for deep links into Teams
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var teamsDeepLinkProtocol = 'https';
/**
 * @hidden
 * The host used for deep links into Teams
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var teamsDeepLinkHost = 'teams.microsoft.com';
/** @hidden */
var errorLibraryNotInitialized = 'The library has not yet been initialized';
/** @hidden */
var errorRuntimeNotInitialized = 'The runtime has not yet been initialized';
/** @hidden */
var errorRuntimeNotSupported = 'The runtime version is not supported';
/** @hidden */
var errorCallNotStarted = 'The call was not properly started';

;// CONCATENATED MODULE: ./src/internal/globalVars.ts
var GlobalVars = /** @class */ (function () {
    function GlobalVars() {
    }
    GlobalVars.initializeCalled = false;
    GlobalVars.initializeCompleted = false;
    GlobalVars.additionalValidOrigins = [];
    GlobalVars.isFramelessWindow = false;
    GlobalVars.printCapabilityEnabled = false;
    return GlobalVars;
}());


// EXTERNAL MODULE: ../../node_modules/.pnpm/debug@4.3.4/node_modules/debug/src/browser.js
var browser = __webpack_require__(302);
;// CONCATENATED MODULE: ./src/internal/telemetry.ts

var topLevelLogger = (0,browser.debug)('teamsJs');
/**
 * @internal
 * Limited to Microsoft-internal use
 *
 * Returns a logger for a given namespace, within the pre-defined top-level teamsJs namespace
 */
function getLogger(namespace) {
    return topLevelLogger.extend(namespace);
}

;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/native.js
const randomUUID = typeof crypto !== 'undefined' && crypto.randomUUID && crypto.randomUUID.bind(crypto);
/* harmony default export */ const esm_browser_native = ({
  randomUUID
});
;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/rng.js
// Unique ID creation requires a high quality random # generator. In the browser we therefore
// require the crypto API and do not support built-in fallback to lower quality random number
// generators (like Math.random()).
let getRandomValues;
const rnds8 = new Uint8Array(16);
function rng() {
  // lazy load so that environments that need to polyfill have a chance to do so
  if (!getRandomValues) {
    // getRandomValues needs to be invoked in a context where "this" is a Crypto implementation.
    getRandomValues = typeof crypto !== 'undefined' && crypto.getRandomValues && crypto.getRandomValues.bind(crypto);

    if (!getRandomValues) {
      throw new Error('crypto.getRandomValues() not supported. See https://github.com/uuidjs/uuid#getrandomvalues-not-supported');
    }
  }

  return getRandomValues(rnds8);
}
;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/stringify.js

/**
 * Convert array of 16 byte values to UUID string format of the form:
 * XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
 */

const byteToHex = [];

for (let i = 0; i < 256; ++i) {
  byteToHex.push((i + 0x100).toString(16).slice(1));
}

function unsafeStringify(arr, offset = 0) {
  // Note: Be careful editing this code!  It's been tuned for performance
  // and works in ways you may not expect. See https://github.com/uuidjs/uuid/pull/434
  return (byteToHex[arr[offset + 0]] + byteToHex[arr[offset + 1]] + byteToHex[arr[offset + 2]] + byteToHex[arr[offset + 3]] + '-' + byteToHex[arr[offset + 4]] + byteToHex[arr[offset + 5]] + '-' + byteToHex[arr[offset + 6]] + byteToHex[arr[offset + 7]] + '-' + byteToHex[arr[offset + 8]] + byteToHex[arr[offset + 9]] + '-' + byteToHex[arr[offset + 10]] + byteToHex[arr[offset + 11]] + byteToHex[arr[offset + 12]] + byteToHex[arr[offset + 13]] + byteToHex[arr[offset + 14]] + byteToHex[arr[offset + 15]]).toLowerCase();
}

function stringify(arr, offset = 0) {
  const uuid = unsafeStringify(arr, offset); // Consistency check for valid UUID.  If this throws, it's likely due to one
  // of the following:
  // - One or more input array values don't map to a hex octet (leading to
  // "undefined" in the uuid)
  // - Invalid input values for the RFC `version` or `variant` fields

  if (!validate(uuid)) {
    throw TypeError('Stringified UUID is invalid');
  }

  return uuid;
}

/* harmony default export */ const esm_browser_stringify = ((/* unused pure expression or super */ null && (stringify)));
;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/v4.js




function v4(options, buf, offset) {
  if (esm_browser_native.randomUUID && !buf && !options) {
    return esm_browser_native.randomUUID();
  }

  options = options || {};
  const rnds = options.random || (options.rng || rng)(); // Per 4.4, set bits for version and `clock_seq_hi_and_reserved`

  rnds[6] = rnds[6] & 0x0f | 0x40;
  rnds[8] = rnds[8] & 0x3f | 0x80; // Copy bytes to buffer, if provided

  if (buf) {
    offset = offset || 0;

    for (let i = 0; i < 16; ++i) {
      buf[offset + i] = rnds[i];
    }

    return buf;
  }

  return unsafeStringify(rnds);
}

/* harmony default export */ const esm_browser_v4 = (v4);
;// CONCATENATED MODULE: ./src/public/interfaces.ts
/* eslint-disable @typescript-eslint/no-explicit-any*/
/**
 * Allowed user file open preferences
 */
var FileOpenPreference;
(function (FileOpenPreference) {
    /** Indicates that the user should be prompted to open the file in inline. */
    FileOpenPreference["Inline"] = "inline";
    /** Indicates that the user should be prompted to open the file in the native desktop application associated with the file type. */
    FileOpenPreference["Desktop"] = "desktop";
    /** Indicates that the user should be prompted to open the file in a web browser. */
    FileOpenPreference["Web"] = "web";
})(FileOpenPreference || (FileOpenPreference = {}));
/**
 * Possible Action Types
 *
 * @beta
 */
var ActionObjectType;
(function (ActionObjectType) {
    /** Represents content within a Microsoft 365 application. */
    ActionObjectType["M365Content"] = "m365content";
})(ActionObjectType || (ActionObjectType = {}));
/**
 * These correspond with field names in the MSGraph.
 * See (commonly accessed resources)[https://learn.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0#commonly-accessed-resources].
 * @beta
 */
var SecondaryM365ContentIdName;
(function (SecondaryM365ContentIdName) {
    /** OneDrive ID */
    SecondaryM365ContentIdName["DriveId"] = "driveId";
    /** Teams Group ID */
    SecondaryM365ContentIdName["GroupId"] = "groupId";
    /** SharePoint ID */
    SecondaryM365ContentIdName["SiteId"] = "siteId";
    /** User ID */
    SecondaryM365ContentIdName["UserId"] = "userId";
})(SecondaryM365ContentIdName || (SecondaryM365ContentIdName = {}));
/** Error codes used to identify different types of errors that can occur while developing apps. */
var ErrorCode;
(function (ErrorCode) {
    /**
     * API not supported in the current platform.
     */
    ErrorCode[ErrorCode["NOT_SUPPORTED_ON_PLATFORM"] = 100] = "NOT_SUPPORTED_ON_PLATFORM";
    /**
     * Internal error encountered while performing the required operation.
     */
    ErrorCode[ErrorCode["INTERNAL_ERROR"] = 500] = "INTERNAL_ERROR";
    /**
     * API is not supported in the current context
     */
    ErrorCode[ErrorCode["NOT_SUPPORTED_IN_CURRENT_CONTEXT"] = 501] = "NOT_SUPPORTED_IN_CURRENT_CONTEXT";
    /**
    Permissions denied by user
    */
    ErrorCode[ErrorCode["PERMISSION_DENIED"] = 1000] = "PERMISSION_DENIED";
    /**
     * Network issue
     */
    ErrorCode[ErrorCode["NETWORK_ERROR"] = 2000] = "NETWORK_ERROR";
    /**
     * Underlying hardware doesn't support the capability
     */
    ErrorCode[ErrorCode["NO_HW_SUPPORT"] = 3000] = "NO_HW_SUPPORT";
    /**
     * One or more arguments are invalid
     */
    ErrorCode[ErrorCode["INVALID_ARGUMENTS"] = 4000] = "INVALID_ARGUMENTS";
    /**
     * User is not authorized for this operation
     */
    ErrorCode[ErrorCode["UNAUTHORIZED_USER_OPERATION"] = 5000] = "UNAUTHORIZED_USER_OPERATION";
    /**
     * Could not complete the operation due to insufficient resources
     */
    ErrorCode[ErrorCode["INSUFFICIENT_RESOURCES"] = 6000] = "INSUFFICIENT_RESOURCES";
    /**
     * Platform throttled the request because of API was invoked too frequently
     */
    ErrorCode[ErrorCode["THROTTLE"] = 7000] = "THROTTLE";
    /**
     * User aborted the operation
     */
    ErrorCode[ErrorCode["USER_ABORT"] = 8000] = "USER_ABORT";
    /**
     * Could not complete the operation in the given time interval
     */
    ErrorCode[ErrorCode["OPERATION_TIMED_OUT"] = 8001] = "OPERATION_TIMED_OUT";
    /**
     * Platform code is old and doesn't implement this API
     */
    ErrorCode[ErrorCode["OLD_PLATFORM"] = 9000] = "OLD_PLATFORM";
    /**
     * The file specified was not found on the given location
     */
    ErrorCode[ErrorCode["FILE_NOT_FOUND"] = 404] = "FILE_NOT_FOUND";
    /**
     * The return value is too big and has exceeded our size boundries
     */
    ErrorCode[ErrorCode["SIZE_EXCEEDED"] = 10000] = "SIZE_EXCEEDED";
})(ErrorCode || (ErrorCode = {}));
/** @hidden */
var DevicePermission;
(function (DevicePermission) {
    DevicePermission["GeoLocation"] = "geolocation";
    DevicePermission["Media"] = "media";
})(DevicePermission || (DevicePermission = {}));

;// CONCATENATED MODULE: ./src/public/constants.ts
/** HostClientType represents the different client platforms on which host can be run. */
var HostClientType;
(function (HostClientType) {
    /** Represents the desktop client of host, which is installed on a user's computer and runs as a standalone application. */
    HostClientType["desktop"] = "desktop";
    /** Represents the web-based client of host, which runs in a web browser. */
    HostClientType["web"] = "web";
    /** Represents the Android mobile client of host, which runs on Android devices such as smartphones and tablets. */
    HostClientType["android"] = "android";
    /** Represents the iOS mobile client of host, which runs on iOS devices such as iPhones. */
    HostClientType["ios"] = "ios";
    /** Represents the iPadOS client of host, which runs on iOS devices such as iPads. */
    HostClientType["ipados"] = "ipados";
    /**
     * @deprecated
     * As of 2.0.0, please use {@link teamsRoomsWindows} instead.
     */
    HostClientType["rigel"] = "rigel";
    /** Represents the client of host, which runs on surface hub devices. */
    HostClientType["surfaceHub"] = "surfaceHub";
    /** Represents the client of host, which runs on Teams Rooms on Windows devices. More information on Microsoft Teams Rooms on Windows can be found [Microsoft Teams Rooms (Windows)](https://support.microsoft.com/office/microsoft-teams-rooms-windows-help-e667f40e-5aab-40c1-bd68-611fe0002ba2)*/
    HostClientType["teamsRoomsWindows"] = "teamsRoomsWindows";
    /** Represents the client of host, which runs on Teams Rooms on Android devices. More information on Microsoft Teams Rooms on Android can be found [Microsoft Teams Rooms (Android)].(https://support.microsoft.com/office/get-started-with-teams-rooms-on-android-68517298-d513-46be-8d6d-d41db5e6b4b2)*/
    HostClientType["teamsRoomsAndroid"] = "teamsRoomsAndroid";
    /** Represents the client of host, which runs on Teams phones. More information can be found [Microsoft Teams Phones](https://support.microsoft.com/office/get-started-with-teams-phones-694ca17d-3ecf-40ca-b45e-d21b2c442412) */
    HostClientType["teamsPhones"] = "teamsPhones";
    /** Represents the client of host, which runs on Teams displays devices. More information can be found [Microsoft Teams Displays](https://support.microsoft.com/office/get-started-with-teams-displays-ff299825-7f13-4528-96c2-1d3437e6d4e6) */
    HostClientType["teamsDisplays"] = "teamsDisplays";
})(HostClientType || (HostClientType = {}));
/** HostName indicates the possible hosts for your application. */
var HostName;
(function (HostName) {
    /**
     * Office.com and Office Windows App
     */
    HostName["office"] = "Office";
    /**
     * For "desktop" specifically, this refers to the new, pre-release version of Outlook for Windows.
     * Also used on other platforms that map to a single Outlook client.
     */
    HostName["outlook"] = "Outlook";
    /**
     * Outlook for Windows: the classic, native, desktop client
     */
    HostName["outlookWin32"] = "OutlookWin32";
    /**
     * Microsoft-internal test Host
     */
    HostName["orange"] = "Orange";
    /**
     * Teams
     */
    HostName["teams"] = "Teams";
    /**
     * Modern Teams
     */
    HostName["teamsModern"] = "TeamsModern";
})(HostName || (HostName = {}));
/**
 * FrameContexts provides information about the context in which the app is running within the host.
 * Developers can use FrameContexts to determine how their app should behave in different contexts,
 * and can use the information provided by the context to adapt the app to the user's needs.
 *
 * @example
 * If your app is running in the "settings" context, you should be displaying your apps configuration page.
 * If the app is running in the content context, the developer may want to display information relevant to
 * the content the user is currently viewing.
 */
var FrameContexts;
(function (FrameContexts) {
    /**
     * App's frame context from where settings page can be accessed.
     * See [how to create a configuration page.]( https://learn.microsoft.com/microsoftteams/platform/tabs/how-to/create-tab-pages/configuration-page?tabs=teamsjs-v2)
     */
    FrameContexts["settings"] = "settings";
    /** The default context for the app where all the content of the app is displayed. */
    FrameContexts["content"] = "content";
    /** Frame context used when app is running in the authentication window launched by calling {@link authentication.authenticate} */
    FrameContexts["authentication"] = "authentication";
    /** The page shown when the user uninstalls the app. */
    FrameContexts["remove"] = "remove";
    /** A task module is a pop-up window that can be used to display a form, a dialog, or other interactive content within the host. */
    FrameContexts["task"] = "task";
    /** The side panel is a persistent panel that is displayed on the right side of the host and can be used to display content or UI that is relevant to the current page or tab. */
    FrameContexts["sidePanel"] = "sidePanel";
    /** The stage is a large area that is displayed at the center of the host and can be used to display content or UI that requires a lot of space, such as a video player or a document editor. */
    FrameContexts["stage"] = "stage";
    /** App's frame context from where meetingStage can be accessed in a meeting session, which is the primary area where video and presentation content is displayed during a meeting. */
    FrameContexts["meetingStage"] = "meetingStage";
})(FrameContexts || (FrameContexts = {}));
/**
 * Indicates the team type, currently used to distinguish between different team
 * types in Office 365 for Education (team types 1, 2, 3, and 4).
 */
var TeamType;
(function (TeamType) {
    /** Represents a standard or classic team in host that is designed for ongoing collaboration and communication among a group of people. */
    TeamType[TeamType["Standard"] = 0] = "Standard";
    /**  Represents an educational team in host that is designed for classroom collaboration and communication among students and teachers. */
    TeamType[TeamType["Edu"] = 1] = "Edu";
    /** Represents a class team in host that is designed for classroom collaboration and communication among students and teachers in a structured environment. */
    TeamType[TeamType["Class"] = 2] = "Class";
    /** Represents a professional learning community (PLC) team in host that is designed for educators to collaborate and share resources and best practices. */
    TeamType[TeamType["Plc"] = 3] = "Plc";
    /** Represents a staff team in host that is designed for staff collaboration and communication among staff members.*/
    TeamType[TeamType["Staff"] = 4] = "Staff";
})(TeamType || (TeamType = {}));
/**
 * Indicates the various types of roles of a user in a team.
 */
var UserTeamRole;
(function (UserTeamRole) {
    /** Represents that the user is an owner or administrator of the team. */
    UserTeamRole[UserTeamRole["Admin"] = 0] = "Admin";
    /** Represents that the user is a standard member of the team. */
    UserTeamRole[UserTeamRole["User"] = 1] = "User";
    /** Represents that the user does not have any role in the team. */
    UserTeamRole[UserTeamRole["Guest"] = 2] = "Guest";
})(UserTeamRole || (UserTeamRole = {}));
/**
 * Dialog module dimension enum
 */
var DialogDimension;
(function (DialogDimension) {
    /** Represents a large-sized dialog box, which is typically used for displaying large amounts of content or complex workflows that require more space. */
    DialogDimension["Large"] = "large";
    /** Represents a medium-sized dialog box, which is typically used for displaying moderate amounts of content or workflows that require less space. */
    DialogDimension["Medium"] = "medium";
    /** Represents a small-sized dialog box, which is typically used for displaying simple messages or workflows that require minimal space.*/
    DialogDimension["Small"] = "small";
})(DialogDimension || (DialogDimension = {}));

/**
 * @deprecated
 * As of 2.0.0, please use {@link DialogDimension} instead.
 */
// eslint-disable-next-line @typescript-eslint/no-unused-vars
var TaskModuleDimension = DialogDimension;
/**
 * The type of the channel with which the content is associated.
 */
var ChannelType;
(function (ChannelType) {
    /** The default channel type. Type of channel is used for general collaboration and communication within a team. */
    ChannelType["Regular"] = "Regular";
    /** Type of channel is used for sensitive or confidential communication within a team and is only accessible to members of the channel. */
    ChannelType["Private"] = "Private";
    /** Type of channel is used for collaboration between multiple teams or groups and is accessible to members of all the teams or groups. */
    ChannelType["Shared"] = "Shared";
})(ChannelType || (ChannelType = {}));
/** An error object indicating that the requested operation or feature is not supported on the current platform or device.
 * @typedef {Object} SdkError
 */
var errorNotSupportedOnPlatform = { errorCode: ErrorCode.NOT_SUPPORTED_ON_PLATFORM };
/**
 * @hidden
 *
 * Minimum Adaptive Card version supported by the host.
 */
var minAdaptiveCardVersion = { majorVersion: 1, minorVersion: 5 };
/**
 * @hidden
 *
 * Adaptive Card version supported by the Teams v1 client.
 */
var teamsMinAdaptiveCardVersion = {
    adaptiveCardSchemaVersion: { majorVersion: 1, minorVersion: 5 },
};

;// CONCATENATED MODULE: ./src/internal/utils.ts
/* eslint-disable @typescript-eslint/ban-types */
/* eslint-disable @typescript-eslint/no-unused-vars */




/**
 * @param pattern - reference pattern
 * @param host - candidate string
 * @returns returns true if host matches pre-know valid pattern
 *
 * @example
 *    validateHostAgainstPattern('*.teams.microsoft.com', 'subdomain.teams.microsoft.com') returns true
 *    validateHostAgainstPattern('teams.microsoft.com', 'team.microsoft.com') returns false
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateHostAgainstPattern(pattern, host) {
    if (pattern.substring(0, 2) === '*.') {
        var suffix = pattern.substring(1);
        if (host.length > suffix.length &&
            host.split('.').length === suffix.split('.').length &&
            host.substring(host.length - suffix.length) === suffix) {
            return true;
        }
    }
    else if (pattern === host) {
        return true;
    }
    return false;
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function validateOrigin(messageOrigin) {
    // Check whether the url is in the pre-known allowlist or supplied by user
    if (!isValidHttpsURL(messageOrigin)) {
        return false;
    }
    var messageOriginHost = messageOrigin.host;
    if (validOrigins.some(function (pattern) { return validateHostAgainstPattern(pattern, messageOriginHost); })) {
        return true;
    }
    for (var _i = 0, _a = GlobalVars.additionalValidOrigins; _i < _a.length; _i++) {
        var domainOrPattern = _a[_i];
        var pattern = domainOrPattern.substring(0, 8) === 'https://' ? domainOrPattern.substring(8) : domainOrPattern;
        if (validateHostAgainstPattern(pattern, messageOriginHost)) {
            return true;
        }
    }
    return false;
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function getGenericOnCompleteHandler(errorMessage) {
    return function (success, reason) {
        if (!success) {
            throw new Error(errorMessage ? errorMessage : reason);
        }
    };
}
/**
 * @hidden
 * Compares SDK versions.
 *
 * @param v1 - first version
 * @param v2 - second version
 * @returns NaN in case inputs are not in right format
 *         -1 if v1 < v2
 *          1 if v1 > v2
 *          0 otherwise
 * @example
 *    compareSDKVersions('1.2', '1.2.0') returns 0
 *    compareSDKVersions('1.2a', '1.2b') returns NaN
 *    compareSDKVersions('1.2', '1.3') returns -1
 *    compareSDKVersions('2.0', '1.3.2') returns 1
 *    compareSDKVersions('2.0', 2.0) returns NaN
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function compareSDKVersions(v1, v2) {
    if (typeof v1 !== 'string' || typeof v2 !== 'string') {
        return NaN;
    }
    var v1parts = v1.split('.');
    var v2parts = v2.split('.');
    function isValidPart(x) {
        // input has to have one or more digits
        // For ex - returns true for '11', false for '1a1', false for 'a', false for '2b'
        return /^\d+$/.test(x);
    }
    if (!v1parts.every(isValidPart) || !v2parts.every(isValidPart)) {
        return NaN;
    }
    // Make length of both parts equal
    while (v1parts.length < v2parts.length) {
        v1parts.push('0');
    }
    while (v2parts.length < v1parts.length) {
        v2parts.push('0');
    }
    for (var i = 0; i < v1parts.length; ++i) {
        if (Number(v1parts[i]) == Number(v2parts[i])) {
            continue;
        }
        else if (Number(v1parts[i]) > Number(v2parts[i])) {
            return 1;
        }
        else {
            return -1;
        }
    }
    return 0;
}
/**
 * @hidden
 * Generates a GUID
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function generateGUID() {
    return esm_browser_v4();
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function utils_deepFreeze(obj) {
    Object.keys(obj).forEach(function (prop) {
        if (typeof obj[prop] === 'object') {
            utils_deepFreeze(obj[prop]);
        }
    });
    return Object.freeze(obj);
}
/**
 * This utility function is used when the result of the promise is same as the result in the callback.
 * @param funcHelper
 * @param callback
 * @param args
 * @returns
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function callCallbackWithErrorOrResultFromPromiseAndReturnPromise(funcHelper, callback) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var args = [];
    for (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var _i = 2; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i < arguments.length; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        args[_i - 2] = arguments[_i];
    }
    var p = funcHelper.apply(void 0, args);
    p.then(function (result) {
        if (callback) {
            callback(undefined, result);
        }
    }).catch(function (e) {
        if (callback) {
            callback(e);
        }
    });
    return p;
}
/**
 * This utility function is used when the return type of the promise is usually void and
 * the result in the callback is a boolean type (true for success and false for error)
 * @param funcHelper
 * @param callback
 * @param args
 * @returns
 * @internal
 * Limited to Microsoft-internal use
 */
function callCallbackWithErrorOrBooleanFromPromiseAndReturnPromise(funcHelper, callback) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var args = [];
    for (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var _i = 2; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i < arguments.length; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        args[_i - 2] = arguments[_i];
    }
    var p = funcHelper.apply(void 0, args);
    p.then(function () {
        if (callback) {
            callback(undefined, true);
        }
    }).catch(function (e) {
        if (callback) {
            callback(e, false);
        }
    });
    return p;
}
/**
 * This utility function is called when the callback has only Error/SdkError as the primary argument.
 * @param funcHelper
 * @param callback
 * @param args
 * @returns
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function callCallbackWithSdkErrorFromPromiseAndReturnPromise(funcHelper, callback) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var args = [];
    for (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var _i = 2; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i < arguments.length; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        args[_i - 2] = arguments[_i];
    }
    var p = funcHelper.apply(void 0, args);
    p.then(function () {
        if (callback) {
            callback(null);
        }
    }).catch(function (e) {
        if (callback) {
            callback(e);
        }
    });
    return p;
}
/**
 * This utility function is used when the result of the promise is same as the result in the callback.
 * @param funcHelper
 * @param callback
 * @param args
 * @returns
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function callCallbackWithErrorOrResultOrNullFromPromiseAndReturnPromise(funcHelper, callback) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var args = [];
    for (
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    var _i = 2; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i < arguments.length; 
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    _i++) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        args[_i - 2] = arguments[_i];
    }
    var p = funcHelper.apply(void 0, args);
    p.then(function (result) {
        if (callback) {
            callback(null, result);
        }
    }).catch(function (e) {
        if (callback) {
            callback(e, null);
        }
    });
    return p;
}
/**
 * A helper function to add a timeout to an asynchronous operation.
 *
 * @param action Action to wrap the timeout around
 * @param timeoutInMs Timeout period in milliseconds
 * @param timeoutError Error to reject the promise with if timeout elapses before the action completed
 * @returns A promise which resolves to the result of provided action or rejects with a provided timeout error
 * if the initial action didn't complete within provided timeout.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function runWithTimeout(action, timeoutInMs, timeoutError) {
    return new Promise(function (resolve, reject) {
        var timeoutHandle = setTimeout(reject, timeoutInMs, timeoutError);
        action()
            .then(function (result) {
            clearTimeout(timeoutHandle);
            resolve(result);
        })
            .catch(function (error) {
            clearTimeout(timeoutHandle);
            reject(error);
        });
    });
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function createTeamsAppLink(params) {
    var url = new URL('https://teams.microsoft.com/l/entity/' +
        encodeURIComponent(params.appId) +
        '/' +
        encodeURIComponent(params.pageId));
    if (params.webUrl) {
        url.searchParams.append('webUrl', params.webUrl);
    }
    if (params.channelId || params.subPageId) {
        url.searchParams.append('context', JSON.stringify({ channelId: params.channelId, subEntityId: params.subPageId }));
    }
    return url.toString();
}
/**
 * @hidden
 * Checks if the Adaptive Card schema version is supported by the host.
 * @param hostAdaptiveCardSchemaVersion Host's supported Adaptive Card version in the runtime.
 *
 * @returns true if the Adaptive Card Version is not supported and false if it is supported.
 */
function isHostAdaptiveCardSchemaVersionUnsupported(hostAdaptiveCardSchemaVersion) {
    var versionCheck = compareSDKVersions("".concat(hostAdaptiveCardSchemaVersion.majorVersion, ".").concat(hostAdaptiveCardSchemaVersion.minorVersion), "".concat(minAdaptiveCardVersion.majorVersion, ".").concat(minAdaptiveCardVersion.minorVersion));
    if (versionCheck >= 0) {
        return false;
    }
    else {
        return true;
    }
}
/**
 * @hidden
 * Checks if a URL is a HTTPS protocol based URL.
 * @param url URL to be validated.
 *
 * @returns true if the URL is an https URL.
 */
function isValidHttpsURL(url) {
    return url.protocol === 'https:';
}

;// CONCATENATED MODULE: ./src/public/runtime.ts
/* eslint-disable @typescript-eslint/ban-types */
var __assign = (undefined && undefined.__assign) || function () {
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





var runtimeLogger = getLogger('runtime');
var latestRuntimeApiVersion = 2;
function isLatestRuntimeVersion(runtime) {
    return runtime.apiVersion === latestRuntimeApiVersion;
}
// Constant used to set the runtime configuration
var _uninitializedRuntime = {
    apiVersion: -1,
    supports: {},
};
/**
 * @hidden
 * Ensures that the runtime has been initialized

 * @returns True if the runtime has been initialized
 * @throws Error if the runtime has not been initialized
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function isRuntimeInitialized(runtime) {
    if (isLatestRuntimeVersion(runtime)) {
        return true;
    }
    else if (runtime.apiVersion === -1) {
        throw new Error(errorRuntimeNotInitialized);
    }
    else {
        throw new Error(errorRuntimeNotSupported);
    }
}
var runtime = _uninitializedRuntime;
var teamsRuntimeConfig = {
    apiVersion: 2,
    hostVersionsInfo: teamsMinAdaptiveCardVersion,
    isLegacyTeams: true,
    supports: {
        appInstallDialog: {},
        appEntity: {},
        call: {},
        chat: {},
        conversations: {},
        dialog: {
            card: {
                bot: {},
            },
            url: {
                bot: {},
            },
            update: {},
        },
        interactive: {},
        logs: {},
        meetingRoom: {},
        menus: {},
        monetization: {},
        notifications: {},
        pages: {
            appButton: {},
            tabs: {},
            config: {},
            backStack: {},
            fullTrust: {},
        },
        remoteCamera: {},
        stageView: {},
        teams: {
            fullTrust: {},
        },
        teamsCore: {},
        video: {
            sharedFrame: {},
        },
    },
};
var v1HostClientTypes = [
    HostClientType.desktop,
    HostClientType.web,
    HostClientType.android,
    HostClientType.ios,
    HostClientType.rigel,
    HostClientType.surfaceHub,
    HostClientType.teamsRoomsWindows,
    HostClientType.teamsRoomsAndroid,
    HostClientType.teamsPhones,
    HostClientType.teamsDisplays,
];
/**
 * @hidden
 * Uses upgradeChain to transform an outdated runtime object to the latest format.
 * @param outdatedRuntime - The runtime object to be upgraded
 * @returns The upgraded runtime object
 * @throws Error if the runtime object could not be upgraded to the latest version
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function fastForwardRuntime(outdatedRuntime) {
    var runtime = outdatedRuntime;
    if (runtime.apiVersion < latestRuntimeApiVersion) {
        upgradeChain.forEach(function (upgrade) {
            if (runtime.apiVersion === upgrade.versionToUpgradeFrom) {
                runtime = upgrade.upgradeToNextVersion(runtime);
            }
        });
    }
    if (isLatestRuntimeVersion(runtime)) {
        return runtime;
    }
    else {
        throw new Error('Received a runtime that could not be upgraded to the latest version');
    }
}
/**
 * @hidden
 * List of transformations required to upgrade a runtime object from a previous version to the next higher version.
 * This list must be ordered from lowest version to highest version
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var upgradeChain = [
    // This upgrade has been included for testing, it may be removed when there is a real upgrade implemented
    {
        versionToUpgradeFrom: 1,
        upgradeToNextVersion: function (previousVersionRuntime) {
            var _a;
            return {
                apiVersion: 2,
                hostVersionsInfo: undefined,
                isLegacyTeams: previousVersionRuntime.isLegacyTeams,
                supports: __assign(__assign({}, previousVersionRuntime.supports), { dialog: previousVersionRuntime.supports.dialog
                        ? {
                            card: undefined,
                            url: previousVersionRuntime.supports.dialog,
                            update: (_a = previousVersionRuntime.supports.dialog) === null || _a === void 0 ? void 0 : _a.update,
                        }
                        : undefined }),
            };
        },
    },
];
var versionConstants = {
    '1.9.0': [
        {
            capability: { location: {} },
            hostClientTypes: v1HostClientTypes,
        },
    ],
    '2.0.0': [
        {
            capability: { people: {} },
            hostClientTypes: v1HostClientTypes,
        },
        {
            capability: { sharing: {} },
            hostClientTypes: [HostClientType.desktop, HostClientType.web],
        },
    ],
    '2.0.1': [
        {
            capability: { teams: { fullTrust: { joinedTeams: {} } } },
            hostClientTypes: [
                HostClientType.android,
                HostClientType.desktop,
                HostClientType.ios,
                HostClientType.teamsRoomsAndroid,
                HostClientType.teamsPhones,
                HostClientType.teamsDisplays,
                HostClientType.web,
            ],
        },
        {
            capability: { webStorage: {} },
            hostClientTypes: [HostClientType.desktop],
        },
    ],
    '2.0.5': [
        {
            capability: { webStorage: {} },
            hostClientTypes: [HostClientType.android, HostClientType.desktop, HostClientType.ios],
        },
    ],
};
var generateBackCompatRuntimeConfigLogger = runtimeLogger.extend('generateBackCompatRuntimeConfig');
/**
 * @internal
 * Limited to Microsoft-internal use
 *
 * Generates and returns a runtime configuration for host clients which are not on the latest host SDK version
 * and do not provide their own runtime config. Their supported capabilities are based on the highest
 * client SDK version that they can support.
 *
 * @param highestSupportedVersion - The highest client SDK version that the host client can support.
 * @returns runtime which describes the APIs supported by the legacy host client.
 */
function generateBackCompatRuntimeConfig(highestSupportedVersion) {
    generateBackCompatRuntimeConfigLogger('generating back compat runtime config for %s', highestSupportedVersion);
    var newSupports = __assign({}, teamsRuntimeConfig.supports);
    generateBackCompatRuntimeConfigLogger('Supported capabilities in config before updating based on highestSupportedVersion: %o', newSupports);
    Object.keys(versionConstants).forEach(function (versionNumber) {
        if (compareSDKVersions(highestSupportedVersion, versionNumber) >= 0) {
            versionConstants[versionNumber].forEach(function (capabilityReqs) {
                if (capabilityReqs.hostClientTypes.includes(GlobalVars.hostClientType)) {
                    newSupports = __assign(__assign({}, newSupports), capabilityReqs.capability);
                }
            });
        }
    });
    var backCompatRuntimeConfig = {
        apiVersion: 2,
        hostVersionsInfo: teamsMinAdaptiveCardVersion,
        isLegacyTeams: true,
        supports: newSupports,
    };
    generateBackCompatRuntimeConfigLogger('Runtime config after updating based on highestSupportedVersion: %o', backCompatRuntimeConfig);
    return backCompatRuntimeConfig;
}
var applyRuntimeConfigLogger = runtimeLogger.extend('applyRuntimeConfig');
function applyRuntimeConfig(runtimeConfig) {
    // Some hosts that have not adopted runtime versioning send a string for apiVersion, so we should handle those as v1 inputs
    if (typeof runtimeConfig.apiVersion === 'string') {
        applyRuntimeConfigLogger('Trying to apply runtime with string apiVersion, processing as v1: %o', runtimeConfig);
        runtimeConfig = __assign(__assign({}, runtimeConfig), { apiVersion: 1 });
    }
    applyRuntimeConfigLogger('Fast-forwarding runtime %o', runtimeConfig);
    var ffRuntimeConfig = fastForwardRuntime(runtimeConfig);
    applyRuntimeConfigLogger('Applying runtime %o', ffRuntimeConfig);
    runtime = utils_deepFreeze(ffRuntimeConfig);
}
function setUnitializedRuntime() {
    runtime = deepFreeze(_uninitializedRuntime);
}
/**
 * @hidden
 * Constant used to set minimum runtime configuration
 * while un-initializing an app in unit test case.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var _minRuntimeConfigToUninitialize = {
    apiVersion: 2,
    supports: {
        pages: {
            appButton: {},
            tabs: {},
            config: {},
            backStack: {},
            fullTrust: {},
        },
        teamsCore: {},
        logs: {},
    },
};

;// CONCATENATED MODULE: ./src/public/version.ts
/**
 * @hidden
 *  Package version.
 */
var version = "2.13.0";

;// CONCATENATED MODULE: ./src/internal/internalAPIs.ts







var internalLogger = getLogger('internal');
var ensureInitializeCalledLogger = internalLogger.extend('ensureInitializeCalled');
var ensureInitializedLogger = internalLogger.extend('ensureInitialized');
/**
 * Ensures `initialize` was called. This function does NOT verify that a response from Host was received and initialization completed.
 *
 * `ensureInitializeCalled` should only be used for APIs which:
 * - work in all FrameContexts
 * - are part of a required Capability
 * - are suspected to be used directly after calling `initialize`, potentially without awaiting the `initialize` call itself
 *
 * For most APIs {@link ensureInitialized} is the right validation function to use instead.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function ensureInitializeCalled() {
    if (!GlobalVars.initializeCalled) {
        ensureInitializeCalledLogger(errorLibraryNotInitialized);
        throw new Error(errorLibraryNotInitialized);
    }
}
/**
 * Ensures `initialize` was called and response from Host was received and processed and that `runtime` is initialized.
 * If expected FrameContexts are provided, it also validates that the current FrameContext matches one of the expected ones.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function ensureInitialized(runtime) {
    var expectedFrameContexts = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        expectedFrameContexts[_i - 1] = arguments[_i];
    }
    // This global var can potentially be removed in the future if we use the initialization status of the runtime object as our source of truth
    if (!GlobalVars.initializeCompleted) {
        ensureInitializedLogger('%s. initializeCalled: %s', errorLibraryNotInitialized, GlobalVars.initializeCalled.toString());
        throw new Error(errorLibraryNotInitialized);
    }
    if (expectedFrameContexts && expectedFrameContexts.length > 0) {
        var found = false;
        for (var i = 0; i < expectedFrameContexts.length; i++) {
            if (expectedFrameContexts[i] === GlobalVars.frameContext) {
                found = true;
                break;
            }
        }
        if (!found) {
            throw new Error("This call is only allowed in following contexts: ".concat(JSON.stringify(expectedFrameContexts), ". ") +
                "Current context: \"".concat(GlobalVars.frameContext, "\"."));
        }
    }
    return isRuntimeInitialized(runtime);
}
/**
 * @hidden
 * Checks whether the platform has knowledge of this API by doing a comparison
 * on API required version and platform supported version of the SDK
 *
 * @param requiredVersion - SDK version required by the API
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function isCurrentSDKVersionAtLeast(requiredVersion) {
    if (requiredVersion === void 0) { requiredVersion = defaultSDKVersionForCompatCheck; }
    var value = compareSDKVersions(GlobalVars.clientSupportedSDKVersion, requiredVersion);
    if (isNaN(value)) {
        return false;
    }
    return value >= 0;
}
/**
 * @hidden
 * Helper function to identify if host client is either android, ios, or ipados
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function isHostClientMobile() {
    return (GlobalVars.hostClientType == HostClientType.android ||
        GlobalVars.hostClientType == HostClientType.ios ||
        GlobalVars.hostClientType == HostClientType.ipados);
}
/**
 * @hidden
 * Helper function which indicates if current API is supported on mobile or not.
 * @throws SdkError if host client is not android/ios or if the requiredVersion is not
 *          supported by platform or not. Null is returned in case of success.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function throwExceptionIfMobileApiIsNotSupported(requiredVersion) {
    if (requiredVersion === void 0) { requiredVersion = defaultSDKVersionForCompatCheck; }
    if (!isHostClientMobile()) {
        var notSupportedError = { errorCode: ErrorCode.NOT_SUPPORTED_ON_PLATFORM };
        throw notSupportedError;
    }
    else if (!isCurrentSDKVersionAtLeast(requiredVersion)) {
        var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
        throw oldPlatformError;
    }
}
/**
 * @hidden
 * Processes the valid origins specifuied by the user, de-duplicates and converts them into a regexp
 * which is used later for message source/origin validation
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function processAdditionalValidOrigins(validMessageOrigins) {
    var combinedOriginUrls = GlobalVars.additionalValidOrigins.concat(validMessageOrigins.filter(function (_origin) {
        return typeof _origin === 'string' && userOriginUrlValidationRegExp.test(_origin);
    }));
    var dedupUrls = {};
    combinedOriginUrls = combinedOriginUrls.filter(function (_originUrl) {
        if (dedupUrls[_originUrl]) {
            return false;
        }
        dedupUrls[_originUrl] = true;
        return true;
    });
    GlobalVars.additionalValidOrigins = combinedOriginUrls;
}

;// CONCATENATED MODULE: ./src/private/inServerSideRenderingEnvironment.ts
function inServerSideRenderingEnvironment() {
    return typeof window === 'undefined';
}

;// CONCATENATED MODULE: ./src/public/authentication.ts






/**
 * Namespace to interact with the authentication-specific part of the SDK.
 *
 * This object is used for starting or completing authentication flows.
 */
var authentication;
(function (authentication) {
    var authHandlers;
    var authWindowMonitor;
    /**
     * @hidden
     * @internal
     * Limited to Microsoft-internal use; automatically called when library is initialized
     */
    function initialize() {
        registerHandler('authentication.authenticate.success', handleSuccess, false);
        registerHandler('authentication.authenticate.failure', handleFailure, false);
    }
    authentication.initialize = initialize;
    var authParams;
    /**
     * @deprecated
     * As of 2.0.0, this function has been deprecated in favor of a Promise-based pattern using {@link authentication.authenticate authentication.authenticate(authenticateParameters: AuthenticatePopUpParameters): Promise\<string\>}
     *
     * Registers handlers to be called with the result of an authentication flow triggered using {@link authentication.authenticate authentication.authenticate(authenticateParameters?: AuthenticateParameters): void}
     *
     * @param authenticateParameters - Configuration for authentication flow pop-up result communication
     */
    function registerAuthenticationHandlers(authenticateParameters) {
        authParams = authenticateParameters;
    }
    authentication.registerAuthenticationHandlers = registerAuthenticationHandlers;
    function authenticate(authenticateParameters) {
        var isDifferentParamsInCall = authenticateParameters !== undefined;
        var authenticateParams = isDifferentParamsInCall
            ? authenticateParameters
            : authParams;
        if (!authenticateParams) {
            throw new Error('No parameters are provided for authentication');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.remove, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
        return authenticateHelper(authenticateParams)
            .then(function (value) {
            try {
                if (authenticateParams && authenticateParams.successCallback) {
                    authenticateParams.successCallback(value);
                    return '';
                }
                return value;
            }
            finally {
                if (!isDifferentParamsInCall) {
                    authParams = null;
                }
            }
        })
            .catch(function (err) {
            try {
                if (authenticateParams && authenticateParams.failureCallback) {
                    authenticateParams.failureCallback(err.message);
                    return '';
                }
                throw err;
            }
            finally {
                if (!isDifferentParamsInCall) {
                    authParams = null;
                }
            }
        });
    }
    authentication.authenticate = authenticate;
    function authenticateHelper(authenticateParameters) {
        return new Promise(function (resolve, reject) {
            if (GlobalVars.hostClientType === HostClientType.desktop ||
                GlobalVars.hostClientType === HostClientType.android ||
                GlobalVars.hostClientType === HostClientType.ios ||
                GlobalVars.hostClientType === HostClientType.ipados ||
                GlobalVars.hostClientType === HostClientType.rigel ||
                GlobalVars.hostClientType === HostClientType.teamsRoomsWindows ||
                GlobalVars.hostClientType === HostClientType.teamsRoomsAndroid ||
                GlobalVars.hostClientType === HostClientType.teamsPhones ||
                GlobalVars.hostClientType === HostClientType.teamsDisplays) {
                // Convert any relative URLs into absolute URLs before sending them over to the parent window.
                var link = document.createElement('a');
                link.href = authenticateParameters.url;
                // Ask the parent window to open an authentication window with the parameters provided by the caller.
                resolve(sendMessageToParentAsync('authentication.authenticate', [
                    link.href,
                    authenticateParameters.width,
                    authenticateParameters.height,
                    authenticateParameters.isExternal,
                ]).then(function (_a) {
                    var success = _a[0], response = _a[1];
                    if (success) {
                        return response;
                    }
                    else {
                        throw new Error(response);
                    }
                }));
            }
            else {
                // Open an authentication window with the parameters provided by the caller.
                authHandlers = {
                    success: resolve,
                    fail: reject,
                };
                openAuthenticationWindow(authenticateParameters);
            }
        });
    }
    function getAuthToken(authTokenRequest) {
        ensureInitializeCalled();
        return getAuthTokenHelper(authTokenRequest)
            .then(function (value) {
            if (authTokenRequest && authTokenRequest.successCallback) {
                authTokenRequest.successCallback(value);
                return '';
            }
            return value;
        })
            .catch(function (err) {
            if (authTokenRequest && authTokenRequest.failureCallback) {
                authTokenRequest.failureCallback(err.message);
                return '';
            }
            throw err;
        });
    }
    authentication.getAuthToken = getAuthToken;
    function getAuthTokenHelper(authTokenRequest) {
        return new Promise(function (resolve) {
            resolve(sendMessageToParentAsync('authentication.getAuthToken', [
                authTokenRequest === null || authTokenRequest === void 0 ? void 0 : authTokenRequest.resources,
                authTokenRequest === null || authTokenRequest === void 0 ? void 0 : authTokenRequest.claims,
                authTokenRequest === null || authTokenRequest === void 0 ? void 0 : authTokenRequest.silent,
            ]));
        }).then(function (_a) {
            var success = _a[0], result = _a[1];
            if (success) {
                return result;
            }
            else {
                throw new Error(result);
            }
        });
    }
    function getUser(userRequest) {
        ensureInitializeCalled();
        return getUserHelper()
            .then(function (value) {
            if (userRequest && userRequest.successCallback) {
                userRequest.successCallback(value);
                return null;
            }
            return value;
        })
            .catch(function (err) {
            if (userRequest && userRequest.failureCallback) {
                userRequest.failureCallback(err.message);
                return null;
            }
            throw err;
        });
    }
    authentication.getUser = getUser;
    function getUserHelper() {
        return new Promise(function (resolve) {
            resolve(sendMessageToParentAsync('authentication.getUser'));
        }).then(function (_a) {
            var success = _a[0], result = _a[1];
            if (success) {
                return result;
            }
            else {
                throw new Error(result);
            }
        });
    }
    function closeAuthenticationWindow() {
        // Stop monitoring the authentication window
        stopAuthenticationWindowMonitor();
        // Try to close the authentication window and clear all properties associated with it
        try {
            if (Communication.childWindow) {
                Communication.childWindow.close();
            }
        }
        finally {
            Communication.childWindow = null;
            Communication.childOrigin = null;
        }
    }
    function openAuthenticationWindow(authenticateParameters) {
        // Close the previously opened window if we have one
        closeAuthenticationWindow();
        // Start with a sensible default size
        var width = authenticateParameters.width || 600;
        var height = authenticateParameters.height || 400;
        // Ensure that the new window is always smaller than our app's window so that it never fully covers up our app
        width = Math.min(width, Communication.currentWindow.outerWidth - 400);
        height = Math.min(height, Communication.currentWindow.outerHeight - 200);
        // Convert any relative URLs into absolute URLs before sending them over to the parent window
        var link = document.createElement('a');
        link.href = authenticateParameters.url.replace('{oauthRedirectMethod}', 'web');
        // We are running in the browser, so we need to center the new window ourselves
        var left = typeof Communication.currentWindow.screenLeft !== 'undefined'
            ? Communication.currentWindow.screenLeft
            : Communication.currentWindow.screenX;
        var top = typeof Communication.currentWindow.screenTop !== 'undefined'
            ? Communication.currentWindow.screenTop
            : Communication.currentWindow.screenY;
        left += Communication.currentWindow.outerWidth / 2 - width / 2;
        top += Communication.currentWindow.outerHeight / 2 - height / 2;
        // Open a child window with a desired set of standard browser features
        Communication.childWindow = Communication.currentWindow.open(link.href, '_blank', 'toolbar=no, location=yes, status=no, menubar=no, scrollbars=yes, top=' +
            top +
            ', left=' +
            left +
            ', width=' +
            width +
            ', height=' +
            height);
        if (Communication.childWindow) {
            // Start monitoring the authentication window so that we can detect if it gets closed before the flow completes
            startAuthenticationWindowMonitor();
        }
        else {
            // If we failed to open the window, fail the authentication flow
            handleFailure('FailedToOpenWindow');
        }
    }
    function stopAuthenticationWindowMonitor() {
        if (authWindowMonitor) {
            clearInterval(authWindowMonitor);
            authWindowMonitor = 0;
        }
        removeHandler('initialize');
        removeHandler('navigateCrossDomain');
    }
    function startAuthenticationWindowMonitor() {
        // Stop the previous window monitor if one is running
        stopAuthenticationWindowMonitor();
        // Create an interval loop that
        // - Notifies the caller of failure if it detects that the authentication window is closed
        // - Keeps pinging the authentication window while it is open to re-establish
        //   contact with any pages along the authentication flow that need to communicate
        //   with us
        authWindowMonitor = Communication.currentWindow.setInterval(function () {
            if (!Communication.childWindow || Communication.childWindow.closed) {
                handleFailure('CancelledByUser');
            }
            else {
                var savedChildOrigin = Communication.childOrigin;
                try {
                    Communication.childOrigin = '*';
                    sendMessageEventToChild('ping');
                }
                finally {
                    Communication.childOrigin = savedChildOrigin;
                }
            }
        }, 100);
        // Set up an initialize-message handler that gives the authentication window its frame context
        registerHandler('initialize', function () {
            return [FrameContexts.authentication, GlobalVars.hostClientType];
        });
        // Set up a navigateCrossDomain message handler that blocks cross-domain re-navigation attempts
        // in the authentication window. We could at some point choose to implement this method via a call to
        // authenticationWindow.location.href = url; however, we would first need to figure out how to
        // validate the URL against the tab's list of valid domains.
        registerHandler('navigateCrossDomain', function () {
            return false;
        });
    }
    /**
     * When using {@link authentication.authenticate authentication.authenticate(authenticateParameters: AuthenticatePopUpParameters): Promise\<string\>}, the
     * window that was opened to execute the authentication flow should call this method after authentiction to notify the caller of
     * {@link authentication.authenticate authentication.authenticate(authenticateParameters: AuthenticatePopUpParameters): Promise\<string\>} that the
     * authentication request was successful.
     *
     * @remarks
     * This function is usable only from the authentication window.
     * This call causes the authentication window to be closed.
     *
     * @param result - Specifies a result for the authentication. If specified, the frame that initiated the authentication pop-up receives
     * this value in its callback or via the `Promise` return value
     * @param callbackUrl - Specifies the url to redirect back to if the client is Win32 Outlook.
     */
    function notifySuccess(result, callbackUrl) {
        redirectIfWin32Outlook(callbackUrl, 'result', result);
        ensureInitialized(runtime, FrameContexts.authentication);
        sendMessageToParent('authentication.authenticate.success', [result]);
        // Wait for the message to be sent before closing the window
        waitForMessageQueue(Communication.parentWindow, function () { return setTimeout(function () { return Communication.currentWindow.close(); }, 200); });
    }
    authentication.notifySuccess = notifySuccess;
    /**
     * When using {@link authentication.authenticate authentication.authenticate(authenticateParameters: AuthenticatePopUpParameters): Promise\<string\>}, the
     * window that was opened to execute the authentication flow should call this method after authentiction to notify the caller of
     * {@link authentication.authenticate authentication.authenticate(authenticateParameters: AuthenticatePopUpParameters): Promise\<string\>} that the
     * authentication request failed.
  
     *
     * @remarks
     * This function is usable only on the authentication window.
     * This call causes the authentication window to be closed.
     *
     * @param result - Specifies a result for the authentication. If specified, the frame that initiated the authentication pop-up receives
     * this value in its callback or via the `Promise` return value
     * @param callbackUrl - Specifies the url to redirect back to if the client is Win32 Outlook.
     */
    function notifyFailure(reason, callbackUrl) {
        redirectIfWin32Outlook(callbackUrl, 'reason', reason);
        ensureInitialized(runtime, FrameContexts.authentication);
        sendMessageToParent('authentication.authenticate.failure', [reason]);
        // Wait for the message to be sent before closing the window
        waitForMessageQueue(Communication.parentWindow, function () { return setTimeout(function () { return Communication.currentWindow.close(); }, 200); });
    }
    authentication.notifyFailure = notifyFailure;
    function handleSuccess(result) {
        try {
            if (authHandlers) {
                authHandlers.success(result);
            }
        }
        finally {
            authHandlers = null;
            closeAuthenticationWindow();
        }
    }
    function handleFailure(reason) {
        try {
            if (authHandlers) {
                authHandlers.fail(new Error(reason));
            }
        }
        finally {
            authHandlers = null;
            closeAuthenticationWindow();
        }
    }
    /**
     * Validates that the callbackUrl param is a valid connector url, appends the result/reason and authSuccess/authFailure as URL fragments and redirects the window
     * @param callbackUrl - the connectors url to redirect to
     * @param key - "result" in case of success and "reason" in case of failure
     * @param value - the value of the passed result/reason parameter
     */
    function redirectIfWin32Outlook(callbackUrl, key, value) {
        if (callbackUrl) {
            var link = document.createElement('a');
            link.href = decodeURIComponent(callbackUrl);
            if (link.host &&
                link.host !== window.location.host &&
                link.host === 'outlook.office.com' &&
                link.search.indexOf('client_type=Win32_Outlook') > -1) {
                if (key && key === 'result') {
                    if (value) {
                        link.href = updateUrlParameter(link.href, 'result', value);
                    }
                    Communication.currentWindow.location.assign(updateUrlParameter(link.href, 'authSuccess', ''));
                }
                if (key && key === 'reason') {
                    if (value) {
                        link.href = updateUrlParameter(link.href, 'reason', value);
                    }
                    Communication.currentWindow.location.assign(updateUrlParameter(link.href, 'authFailure', ''));
                }
            }
        }
    }
    /**
     * Appends either result or reason as a fragment to the 'callbackUrl'
     * @param uri - the url to modify
     * @param key - the fragment key
     * @param value - the fragment value
     */
    function updateUrlParameter(uri, key, value) {
        var i = uri.indexOf('#');
        var hash = i === -1 ? '#' : uri.substr(i);
        hash = hash + '&' + key + (value !== '' ? '=' + value : '');
        uri = i === -1 ? uri : uri.substr(0, i);
        return uri + hash;
    }
    /**
     * @hidden
     * Limited set of data residencies information exposed to 1P application developers
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var DataResidency;
    (function (DataResidency) {
        /**
         * Public
         */
        DataResidency["Public"] = "public";
        /**
         * European Union Data Boundary
         */
        DataResidency["EUDB"] = "eudb";
        /**
         * Other, stored to cover fields that will not be exposed
         */
        DataResidency["Other"] = "other";
    })(DataResidency = authentication.DataResidency || (authentication.DataResidency = {}));
})(authentication || (authentication = {}));

;// CONCATENATED MODULE: ./src/public/dialog.ts
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/ban-types */
/* eslint-disable @typescript-eslint/no-unused-vars */







/**
 * This group of capabilities enables apps to show modal dialogs. There are two primary types of dialogs: URL-based dialogs and [Adaptive Card](https://learn.microsoft.com/adaptive-cards/) dialogs.
 * Both types of dialogs are shown on top of your app, preventing interaction with your app while they are displayed.
 * - URL-based dialogs allow you to specify a URL from which the contents will be shown inside the dialog.
 *   - For URL dialogs, use the functions and interfaces in the {@link dialog.url} namespace.
 * - Adaptive Card-based dialogs allow you to provide JSON describing an Adaptive Card that will be shown inside the dialog.
 *   - For Adaptive Card dialogs, use the functions and interfaces in the {@link dialog.adaptiveCard} namespace.
 *
 * @remarks Note that dialogs were previously called "task modules". While they have been renamed for clarity, the functionality has been maintained.
 * For more details, see [Dialogs](https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/what-are-task-modules)
 *
 * @beta
 */
var dialog;
(function (dialog) {
    var storedMessages = [];
    /**
     * @hidden
     * Hide from docs because this function is only used during initialization
     *
     * Adds register handlers for messageForChild upon initialization and only in the tasks FrameContext. {@link FrameContexts.task}
     * Function is called during app initialization
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    function initialize() {
        registerHandler('messageForChild', handleDialogMessage, false);
    }
    dialog.initialize = initialize;
    function handleDialogMessage(message) {
        if (!GlobalVars.frameContext) {
            // GlobalVars.frameContext is currently not set
            return;
        }
        if (GlobalVars.frameContext === FrameContexts.task) {
            storedMessages.push(message);
        }
        else {
            // Not in task FrameContext, remove 'messageForChild' handler
            removeHandler('messageForChild');
        }
    }
    var url;
    (function (url) {
        /**
         * Allows app to open a url based dialog.
         *
         * @remarks
         * This function cannot be called from inside of a dialog
         *
         * @param urlDialogInfo - An object containing the parameters of the dialog module.
         * @param submitHandler - Handler that triggers when a dialog calls the {@linkcode submit} function or when the user closes the dialog.
         * @param messageFromChildHandler - Handler that triggers if dialog sends a message to the app.
         *
         * @beta
         */
        function open(urlDialogInfo, submitHandler, messageFromChildHandler) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (messageFromChildHandler) {
                registerHandler('messageForParent', messageFromChildHandler);
            }
            var dialogInfo = getDialogInfoFromUrlDialogInfo(urlDialogInfo);
            sendMessageToParent('tasks.startTask', [dialogInfo], function (err, result) {
                submitHandler === null || submitHandler === void 0 ? void 0 : submitHandler({ err: err, result: result });
                removeHandler('messageForParent');
            });
        }
        url.open = open;
        /**
         * Submit the dialog module and close the dialog
         *
         * @remarks
         * This function is only intended to be called from code running within the dialog. Calling it from outside the dialog will have no effect.
         *
         * @param result - The result to be sent to the bot or the app. Typically a JSON object or a serialized version of it,
         *  If this function is called from a dialog while {@link M365ContentAction} is set in the context object by the host, result will be ignored
         *
         * @param appIds - Valid application(s) that can receive the result of the submitted dialogs. Specifying this parameter helps prevent malicious apps from retrieving the dialog result. Multiple app IDs can be specified because a web app from a single underlying domain can power multiple apps across different environments and branding schemes.
         *
         * @beta
         */
        function submit(result, appIds) {
            // FrameContext content should not be here because dialog.submit can be called only from inside of a dialog (FrameContext task)
            // but it's here because Teams mobile incorrectly returns FrameContext.content when calling app.getFrameContext().
            // FrameContexts.content will be removed once the bug is fixed.
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            // Send tasks.completeTask instead of tasks.submitTask message for backward compatibility with Mobile clients
            sendMessageToParent('tasks.completeTask', [result, appIds ? (Array.isArray(appIds) ? appIds : [appIds]) : []]);
        }
        url.submit = submit;
        /**
         *  Send message to the parent from dialog
         *
         * @remarks
         * This function is only intended to be called from code running within the dialog. Calling it from outside the dialog will have no effect.
         *
         * @param message - The message to send to the parent
         *
         * @beta
         */
        function sendMessageToParentFromDialog(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        message) {
            ensureInitialized(runtime, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('messageForParent', [message]);
        }
        url.sendMessageToParentFromDialog = sendMessageToParentFromDialog;
        /**
         *  Send message to the dialog from the parent
         *
         * @param message - The message to send
         *
         * @beta
         */
        function sendMessageToDialog(
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        message) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('messageForChild', [message]);
        }
        url.sendMessageToDialog = sendMessageToDialog;
        /**
         * Register a listener that will be triggered when a message is received from the app that opened the dialog.
         *
         * @remarks
         * This function is only intended to be called from code running within the dialog. Calling it from outside the dialog will have no effect.
         *
         * @param listener - The listener that will be triggered.
         *
         * @beta
         */
        function registerOnMessageFromParent(listener) {
            ensureInitialized(runtime, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            // We need to remove the original 'messageForChild'
            // handler since the original does not allow for post messages.
            // It is replaced by the user specified listener that is passed in.
            removeHandler('messageForChild');
            registerHandler('messageForChild', listener);
            storedMessages.reverse();
            while (storedMessages.length > 0) {
                var message = storedMessages.pop();
                listener(message);
            }
        }
        url.registerOnMessageFromParent = registerOnMessageFromParent;
        /**
         * Checks if dialog.url module is supported by the host
         *
         * @returns boolean to represent whether dialog.url module is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @beta
         */
        function isSupported() {
            return ensureInitialized(runtime) && (runtime.supports.dialog && runtime.supports.dialog.url) !== undefined;
        }
        url.isSupported = isSupported;
        /**
         * Namespace to open a dialog that sends results to the bot framework
         *
         * @beta
         */
        var bot;
        (function (bot) {
            /**
             * Allows an app to open the dialog module using bot.
             *
             * @param botUrlDialogInfo - An object containing the parameters of the dialog module including completionBotId.
             * @param submitHandler - Handler that triggers when the dialog has been submitted or closed.
             * @param messageFromChildHandler - Handler that triggers if dialog sends a message to the app.
             *
             * @returns a function that can be used to send messages to the dialog.
             *
             * @beta
             */
            function open(botUrlDialogInfo, submitHandler, messageFromChildHandler) {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                if (messageFromChildHandler) {
                    registerHandler('messageForParent', messageFromChildHandler);
                }
                var dialogInfo = getDialogInfoFromBotUrlDialogInfo(botUrlDialogInfo);
                sendMessageToParent('tasks.startTask', [dialogInfo], function (err, result) {
                    submitHandler === null || submitHandler === void 0 ? void 0 : submitHandler({ err: err, result: result });
                    removeHandler('messageForParent');
                });
            }
            bot.open = open;
            /**
             * Checks if dialog.url.bot capability is supported by the host
             *
             * @returns boolean to represent whether dialog.url.bot is supported
             *
             * @throws Error if {@linkcode app.initialize} has not successfully completed
             *
             * @beta
             */
            function isSupported() {
                return (ensureInitialized(runtime) &&
                    (runtime.supports.dialog && runtime.supports.dialog.url && runtime.supports.dialog.url.bot) !== undefined);
            }
            bot.isSupported = isSupported;
        })(bot = url.bot || (url.bot = {}));
        /**
         * @hidden
         *
         * Convert UrlDialogInfo to DialogInfo to send the information to host in {@linkcode open} API.
         *
         * @internal
         * Limited to Microsoft-internal use
         */
        function getDialogInfoFromUrlDialogInfo(urlDialogInfo) {
            var dialogInfo = {
                url: urlDialogInfo.url,
                height: urlDialogInfo.size ? urlDialogInfo.size.height : DialogDimension.Small,
                width: urlDialogInfo.size ? urlDialogInfo.size.width : DialogDimension.Small,
                title: urlDialogInfo.title,
                fallbackUrl: urlDialogInfo.fallbackUrl,
            };
            return dialogInfo;
        }
        url.getDialogInfoFromUrlDialogInfo = getDialogInfoFromUrlDialogInfo;
        /**
         * @hidden
         *
         * Convert BotUrlDialogInfo to DialogInfo to send the information to host in {@linkcode bot.open} API.
         *
         * @internal
         * Limited to Microsoft-internal use
         */
        function getDialogInfoFromBotUrlDialogInfo(botUrlDialogInfo) {
            var dialogInfo = getDialogInfoFromUrlDialogInfo(botUrlDialogInfo);
            dialogInfo.completionBotId = botUrlDialogInfo.completionBotId;
            return dialogInfo;
        }
        url.getDialogInfoFromBotUrlDialogInfo = getDialogInfoFromBotUrlDialogInfo;
    })(url = dialog.url || (dialog.url = {}));
    /**
     * This function currently serves no purpose and should not be used. All functionality that used
     * to be covered by this method is now in subcapabilities and those isSupported methods should be
     * used directly.
     *
     * @hidden
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.dialog ? true : false;
    }
    dialog.isSupported = isSupported;
    /**
     * Namespace to update the dialog
     *
     * @beta
     */
    var update;
    (function (update) {
        /**
         * Update dimensions - height/width of a dialog.
         *
         * @param dimensions - An object containing width and height properties.
         *
         * @beta
         */
        function resize(dimensions) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.task, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('tasks.updateTask', [dimensions]);
        }
        update.resize = resize;
        /**
         * Checks if dialog.update capability is supported by the host
         * @returns boolean to represent whether dialog.update capabilty is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @beta
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.dialog
                ? runtime.supports.dialog.update
                    ? true
                    : false
                : false;
        }
        update.isSupported = isSupported;
    })(update = dialog.update || (dialog.update = {}));
    /**
     * Subcapability for interacting with adaptive card dialogs
     * @beta
     */
    var adaptiveCard;
    (function (adaptiveCard) {
        /**
         * Allows app to open an adaptive card based dialog.
         *
         * @remarks
         * This function cannot be called from inside of a dialog
         *
         * @param adaptiveCardDialogInfo - An object containing the parameters of the dialog module {@link AdaptiveCardDialogInfo}.
         * @param submitHandler - Handler that triggers when a dialog calls the {@linkcode url.submit} function or when the user closes the dialog.
         *
         * @beta
         */
        function open(adaptiveCardDialogInfo, submitHandler) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            var dialogInfo = getDialogInfoFromAdaptiveCardDialogInfo(adaptiveCardDialogInfo);
            sendMessageToParent('tasks.startTask', [dialogInfo], function (err, result) {
                submitHandler === null || submitHandler === void 0 ? void 0 : submitHandler({ err: err, result: result });
            });
        }
        adaptiveCard.open = open;
        /**
         * Checks if dialog.adaptiveCard module is supported by the host
         *
         * @returns boolean to represent whether dialog.adaptiveCard module is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @beta
         */
        function isSupported() {
            var isAdaptiveCardVersionSupported = runtime.hostVersionsInfo &&
                runtime.hostVersionsInfo.adaptiveCardSchemaVersion &&
                !isHostAdaptiveCardSchemaVersionUnsupported(runtime.hostVersionsInfo.adaptiveCardSchemaVersion);
            return (ensureInitialized(runtime) &&
                (isAdaptiveCardVersionSupported && runtime.supports.dialog && runtime.supports.dialog.card) !== undefined);
        }
        adaptiveCard.isSupported = isSupported;
        /**
         * Namespace for interaction with adaptive card dialogs that need to communicate with the bot framework
         *
         * @beta
         */
        var bot;
        (function (bot) {
            /**
             * Allows an app to open an adaptive card-based dialog module using bot.
             *
             * @param botAdaptiveCardDialogInfo - An object containing the parameters of the dialog module including completionBotId.
             * @param submitHandler - Handler that triggers when the dialog has been submitted or closed.
             *
             * @beta
             */
            function open(botAdaptiveCardDialogInfo, submitHandler) {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                var dialogInfo = getDialogInfoFromBotAdaptiveCardDialogInfo(botAdaptiveCardDialogInfo);
                sendMessageToParent('tasks.startTask', [dialogInfo], function (err, result) {
                    submitHandler === null || submitHandler === void 0 ? void 0 : submitHandler({ err: err, result: result });
                });
            }
            bot.open = open;
            /**
             * Checks if dialog.adaptiveCard.bot capability is supported by the host
             *
             * @returns boolean to represent whether dialog.adaptiveCard.bot is supported
             *
             * @throws Error if {@linkcode app.initialize} has not successfully completed
             *
             * @beta
             */
            function isSupported() {
                var isAdaptiveCardVersionSupported = runtime.hostVersionsInfo &&
                    runtime.hostVersionsInfo.adaptiveCardSchemaVersion &&
                    !isHostAdaptiveCardSchemaVersionUnsupported(runtime.hostVersionsInfo.adaptiveCardSchemaVersion);
                return (ensureInitialized(runtime) &&
                    (isAdaptiveCardVersionSupported &&
                        runtime.supports.dialog &&
                        runtime.supports.dialog.card &&
                        runtime.supports.dialog.card.bot) !== undefined);
            }
            bot.isSupported = isSupported;
        })(bot = adaptiveCard.bot || (adaptiveCard.bot = {}));
        /**
         * @hidden
         * Hide from docs
         * --------
         * Convert AdaptiveCardDialogInfo to DialogInfo to send the information to host in {@linkcode adaptiveCard.open} API.
         *
         * @internal
         */
        function getDialogInfoFromAdaptiveCardDialogInfo(adaptiveCardDialogInfo) {
            var dialogInfo = {
                card: adaptiveCardDialogInfo.card,
                height: adaptiveCardDialogInfo.size ? adaptiveCardDialogInfo.size.height : DialogDimension.Small,
                width: adaptiveCardDialogInfo.size ? adaptiveCardDialogInfo.size.width : DialogDimension.Small,
                title: adaptiveCardDialogInfo.title,
            };
            return dialogInfo;
        }
        /**
         * @hidden
         * Hide from docs
         * --------
         * Convert BotAdaptiveCardDialogInfo to DialogInfo to send the information to host in {@linkcode adaptiveCard.open} API.
         *
         * @internal
         */
        function getDialogInfoFromBotAdaptiveCardDialogInfo(botAdaptiveCardDialogInfo) {
            var dialogInfo = getDialogInfoFromAdaptiveCardDialogInfo(botAdaptiveCardDialogInfo);
            dialogInfo.completionBotId = botAdaptiveCardDialogInfo.completionBotId;
            return dialogInfo;
        }
    })(adaptiveCard = dialog.adaptiveCard || (dialog.adaptiveCard = {}));
})(dialog || (dialog = {}));

;// CONCATENATED MODULE: ./src/public/menus.ts





/**
 * Namespace to interact with the menu-specific part of the SDK.
 * This object is used to show View Configuration, Action Menu and Navigation Bar Menu.
 */
var menus;
(function (menus) {
    /**
     * Defines how a menu item should appear in the NavBar.
     */
    var DisplayMode;
    (function (DisplayMode) {
        /**
         * Only place this item in the NavBar if there's room for it.
         * If there's no room, item is shown in the overflow menu.
         */
        DisplayMode[DisplayMode["ifRoom"] = 0] = "ifRoom";
        /**
         * Never place this item in the NavBar.
         * The item would always be shown in NavBar's overflow menu.
         */
        DisplayMode[DisplayMode["overflowOnly"] = 1] = "overflowOnly";
    })(DisplayMode = menus.DisplayMode || (menus.DisplayMode = {}));
    /**
     * @hidden
     * Represents information about menu item for Action Menu and Navigation Bar Menu.
     */
    var MenuItem = /** @class */ (function () {
        function MenuItem() {
            /**
             * @hidden
             * State of the menu item
             */
            this.enabled = true;
            /**
             * @hidden
             * Whether the menu item is selected or not
             */
            this.selected = false;
        }
        return MenuItem;
    }());
    menus.MenuItem = MenuItem;
    /**
     * @hidden
     * Represents information about type of list to display in Navigation Bar Menu.
     */
    var MenuListType;
    (function (MenuListType) {
        MenuListType["dropDown"] = "dropDown";
        MenuListType["popOver"] = "popOver";
    })(MenuListType = menus.MenuListType || (menus.MenuListType = {}));
    var navBarMenuItemPressHandler;
    var actionMenuItemPressHandler;
    var viewConfigItemPressHandler;
    /**
     * @hidden
     * Register navBarMenuItemPress, actionMenuItemPress, setModuleView handlers.
     *
     * @internal
     * Limited to Microsoft-internal use.
     */
    function initialize() {
        registerHandler('navBarMenuItemPress', handleNavBarMenuItemPress, false);
        registerHandler('actionMenuItemPress', handleActionMenuItemPress, false);
        registerHandler('setModuleView', handleViewConfigItemPress, false);
    }
    menus.initialize = initialize;
    /**
     * @hidden
     * Registers list of view configurations and it's handler.
     * Handler is responsible for listening selection of View Configuration.
     *
     * @param viewConfig - List of view configurations. Minimum 1 value is required.
     * @param handler - The handler to invoke when the user selects view configuration.
     */
    function setUpViews(viewConfig, handler) {
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        viewConfigItemPressHandler = handler;
        sendMessageToParent('setUpViews', [viewConfig]);
    }
    menus.setUpViews = setUpViews;
    function handleViewConfigItemPress(id) {
        if (!viewConfigItemPressHandler || !viewConfigItemPressHandler(id)) {
            ensureInitialized(runtime);
            sendMessageToParent('viewConfigItemPress', [id]);
        }
    }
    /**
     * @hidden
     * Used to set menu items on the Navigation Bar. If icon is available, icon will be shown, otherwise title will be shown.
     *
     * @param items List of MenuItems for Navigation Bar Menu.
     * @param handler The handler to invoke when the user selects menu item.
     */
    function setNavBarMenu(items, handler) {
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        navBarMenuItemPressHandler = handler;
        sendMessageToParent('setNavBarMenu', [items]);
    }
    menus.setNavBarMenu = setNavBarMenu;
    function handleNavBarMenuItemPress(id) {
        if (!navBarMenuItemPressHandler || !navBarMenuItemPressHandler(id)) {
            ensureInitialized(runtime);
            sendMessageToParent('handleNavBarMenuItemPress', [id]);
        }
    }
    /**
     * @hidden
     * Used to show Action Menu.
     *
     * @param params - Parameters for Menu Parameters
     * @param handler - The handler to invoke when the user selects menu item.
     */
    function showActionMenu(params, handler) {
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        actionMenuItemPressHandler = handler;
        sendMessageToParent('showActionMenu', [params]);
    }
    menus.showActionMenu = showActionMenu;
    function handleActionMenuItemPress(id) {
        if (!actionMenuItemPressHandler || !actionMenuItemPressHandler(id)) {
            ensureInitialized(runtime);
            sendMessageToParent('handleActionMenuItemPress', [id]);
        }
    }
    /**
     * Checks if the menus capability is supported by the host
     * @returns boolean to represent whether the menus capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.menus ? true : false;
    }
    menus.isSupported = isSupported;
})(menus || (menus = {}));

;// CONCATENATED MODULE: ./src/public/teamsAPIs.ts

 // Conflict with some names



/**
 * Namespace containing the set of APIs that support Teams-specific functionalities.
 */
var teamsCore;
(function (teamsCore) {
    /**
     * Enable print capability to support printing page using Ctrl+P and cmd+P
     */
    function enablePrintCapability() {
        if (!GlobalVars.printCapabilityEnabled) {
            ensureInitialized(runtime);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            GlobalVars.printCapabilityEnabled = true;
            // adding ctrl+P and cmd+P handler
            document.addEventListener('keydown', function (event) {
                if ((event.ctrlKey || event.metaKey) && event.keyCode === 80) {
                    print();
                    event.cancelBubble = true;
                    event.preventDefault();
                    event.stopImmediatePropagation();
                }
            });
        }
    }
    teamsCore.enablePrintCapability = enablePrintCapability;
    /**
     * default print handler
     */
    function print() {
        if (typeof window !== 'undefined') {
            window.print();
        }
        else {
            // This codepath only exists to enable compilation in a server-side redered environment. In standard usage, the window object should never be undefined so this code path should never run.
            // If this error has actually been thrown, something has gone very wrong and it is a bug
            throw new Error('window object undefined at print call');
        }
    }
    teamsCore.print = print;
    /**
     * Registers a handler to be called when the page has been requested to load.
     *
     * @remarks Check out [App Caching in Teams](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/build-tabs-for-meeting?tabs=desktop%2Cmeeting-chat-view-desktop%2Cmeeting-stage-view-desktop%2Cchannel-meeting-desktop#app-caching)
     * for a more detailed explanation about using this API.
     *
     * @param handler - The handler to invoke when the page is loaded.
     *
     * @beta
     */
    function registerOnLoadHandler(handler) {
        registerOnLoadHandlerHelper(handler, function () {
            if (handler && !isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
        });
    }
    teamsCore.registerOnLoadHandler = registerOnLoadHandler;
    /**
     * @hidden
     * Undocumented helper function with shared code between deprecated version and current version of the registerOnLoadHandler API.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @param handler - The handler to invoke when the page is loaded.
     * @param versionSpecificHelper - The helper function containing logic pertaining to a specific version of the API.
     */
    function registerOnLoadHandlerHelper(handler, versionSpecificHelper) {
        // allow for registration cleanup even when not finished initializing
        handler && ensureInitialized(runtime);
        if (handler && versionSpecificHelper) {
            versionSpecificHelper();
        }
        handlers_registerOnLoadHandler(handler);
    }
    teamsCore.registerOnLoadHandlerHelper = registerOnLoadHandlerHelper;
    /**
     * Registers a handler to be called before the page is unloaded.
     *
     * @remarks Check out [App Caching in Teams](https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/build-tabs-for-meeting?tabs=desktop%2Cmeeting-chat-view-desktop%2Cmeeting-stage-view-desktop%2Cchannel-meeting-desktop#app-caching)
     * for a more detailed explanation about using this API.
     *
     * @param handler - The handler to invoke before the page is unloaded. If this handler returns true the page should
     * invoke the readyToUnload function provided to it once it's ready to be unloaded.
     *
     * @beta
     */
    function registerBeforeUnloadHandler(handler) {
        registerBeforeUnloadHandlerHelper(handler, function () {
            if (handler && !isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
        });
    }
    teamsCore.registerBeforeUnloadHandler = registerBeforeUnloadHandler;
    /**
     * @hidden
     * Undocumented helper function with shared code between deprecated version and current version of the registerBeforeUnloadHandler API.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @param handler - - The handler to invoke before the page is unloaded. If this handler returns true the page should
     * invoke the readyToUnload function provided to it once it's ready to be unloaded.
     * @param versionSpecificHelper - The helper function containing logic pertaining to a specific version of the API.
     */
    function registerBeforeUnloadHandlerHelper(handler, versionSpecificHelper) {
        // allow for registration cleanup even when not finished initializing
        handler && ensureInitialized(runtime);
        if (handler && versionSpecificHelper) {
            versionSpecificHelper();
        }
        handlers_registerBeforeUnloadHandler(handler);
    }
    teamsCore.registerBeforeUnloadHandlerHelper = registerBeforeUnloadHandlerHelper;
    /**
     * Checks if teamsCore capability is supported by the host
     *
     * @returns boolean to represent whether the teamsCore capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.teamsCore ? true : false;
    }
    teamsCore.isSupported = isSupported;
})(teamsCore || (teamsCore = {}));

;// CONCATENATED MODULE: ./src/public/app.ts
/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/no-explicit-any */



 // Conflict with some names













/**
 * Namespace to interact with app initialization and lifecycle.
 */
var app;
(function (app) {
    var appLogger = getLogger('app');
    // ::::::::::::::::::::::: MicrosoftTeams client SDK public API ::::::::::::::::::::
    /** App Initialization Messages */
    app.Messages = {
        /** App loaded. */
        AppLoaded: 'appInitialization.appLoaded',
        /** App initialized successfully. */
        Success: 'appInitialization.success',
        /** App initialization failed. */
        Failure: 'appInitialization.failure',
        /** App initialization expected failure. */
        ExpectedFailure: 'appInitialization.expectedFailure',
    };
    /**
     * Describes errors that caused app initialization to fail
     */
    var FailedReason;
    (function (FailedReason) {
        /**
         * Authentication failed
         */
        FailedReason["AuthFailed"] = "AuthFailed";
        /**
         * The application timed out
         */
        FailedReason["Timeout"] = "Timeout";
        /**
         * The app failed for a different reason
         */
        FailedReason["Other"] = "Other";
    })(FailedReason = app.FailedReason || (app.FailedReason = {}));
    /**
     * Describes expected errors that occurred during an otherwise successful
     * app initialization
     */
    var ExpectedFailureReason;
    (function (ExpectedFailureReason) {
        /**
         * There was a permission error
         */
        ExpectedFailureReason["PermissionError"] = "PermissionError";
        /**
         * The item was not found
         */
        ExpectedFailureReason["NotFound"] = "NotFound";
        /**
         * The network is currently throttled
         */
        ExpectedFailureReason["Throttling"] = "Throttling";
        /**
         * The application is currently offline
         */
        ExpectedFailureReason["Offline"] = "Offline";
        /**
         * The app failed for a different reason
         */
        ExpectedFailureReason["Other"] = "Other";
    })(ExpectedFailureReason = app.ExpectedFailureReason || (app.ExpectedFailureReason = {}));
    /**
     * Checks whether the Teams client SDK has been initialized.
     * @returns whether the Teams client SDK has been initialized.
     */
    function isInitialized() {
        return GlobalVars.initializeCompleted;
    }
    app.isInitialized = isInitialized;
    /**
     * Gets the Frame Context that the App is running in. See {@link FrameContexts} for the list of possible values.
     * @returns the Frame Context.
     */
    function getFrameContext() {
        return GlobalVars.frameContext;
    }
    app.getFrameContext = getFrameContext;
    /**
     * Number of milliseconds we'll give the initialization call to return before timing it out
     */
    var initializationTimeoutInMs = 5000;
    /**
     * Initializes the library.
     *
     * @remarks
     * Initialize must have completed successfully (as determined by the resolved Promise) before any other library calls are made
     *
     * @param validMessageOrigins - Optionally specify a list of cross frame message origins. They must have
     * https: protocol otherwise they will be ignored. Example: https://www.example.com
     * @returns Promise that will be fulfilled when initialization has completed, or rejected if the initialization fails or times out
     */
    function initialize(validMessageOrigins) {
        if (!inServerSideRenderingEnvironment()) {
            return runWithTimeout(function () { return initializeHelper(validMessageOrigins); }, initializationTimeoutInMs, new Error('SDK initialization timed out.'));
        }
        else {
            var initializeLogger = appLogger.extend('initialize');
            // This log statement should NEVER actually be written. This code path exists only to enable compilation in server-side rendering environments.
            // If you EVER see this statement in ANY log file, something has gone horribly wrong and a bug needs to be filed.
            initializeLogger('window object undefined at initialization');
            return Promise.resolve();
        }
    }
    app.initialize = initialize;
    var initializeHelperLogger = appLogger.extend('initializeHelper');
    function initializeHelper(validMessageOrigins) {
        return new Promise(function (resolve) {
            // Independent components might not know whether the SDK is initialized so might call it to be safe.
            // Just no-op if that happens to make it easier to use.
            if (!GlobalVars.initializeCalled) {
                GlobalVars.initializeCalled = true;
                initializeHandlers();
                GlobalVars.initializePromise = initializeCommunication(validMessageOrigins).then(function (_a) {
                    var context = _a.context, clientType = _a.clientType, runtimeConfig = _a.runtimeConfig, _b = _a.clientSupportedSDKVersion, clientSupportedSDKVersion = _b === void 0 ? defaultSDKVersionForCompatCheck : _b;
                    GlobalVars.frameContext = context;
                    GlobalVars.hostClientType = clientType;
                    GlobalVars.clientSupportedSDKVersion = clientSupportedSDKVersion;
                    // Temporary workaround while the Host is updated with the new argument order.
                    // For now, we might receive any of these possibilities:
                    // - `runtimeConfig` in `runtimeConfig` and `clientSupportedSDKVersion` in `clientSupportedSDKVersion`.
                    // - `runtimeConfig` in `clientSupportedSDKVersion` and `clientSupportedSDKVersion` in `runtimeConfig`.
                    // - `clientSupportedSDKVersion` in `runtimeConfig` and no `clientSupportedSDKVersion`.
                    // This code supports any of these possibilities
                    // Teams AppHost won't provide this runtime config
                    // so we assume that if we don't have it, we must be running in Teams.
                    // After Teams updates its client code, we can remove this default code.
                    try {
                        initializeHelperLogger('Parsing %s', runtimeConfig);
                        var givenRuntimeConfig = JSON.parse(runtimeConfig);
                        initializeHelperLogger('Checking if %o is a valid runtime object', givenRuntimeConfig !== null && givenRuntimeConfig !== void 0 ? givenRuntimeConfig : 'null');
                        // Check that givenRuntimeConfig is a valid instance of IBaseRuntime
                        if (!givenRuntimeConfig || !givenRuntimeConfig.apiVersion) {
                            throw new Error('Received runtime config is invalid');
                        }
                        runtimeConfig && applyRuntimeConfig(givenRuntimeConfig);
                    }
                    catch (e) {
                        if (e instanceof SyntaxError) {
                            try {
                                initializeHelperLogger('Attempting to parse %s as an SDK version', runtimeConfig);
                                // if the given runtime config was actually meant to be a SDK version, store it as such.
                                // TODO: This is a temporary workaround to allow Teams to store clientSupportedSDKVersion even when
                                // it doesn't provide the runtimeConfig. After Teams updates its client code, we should
                                // remove this feature.
                                if (!isNaN(compareSDKVersions(runtimeConfig, defaultSDKVersionForCompatCheck))) {
                                    GlobalVars.clientSupportedSDKVersion = runtimeConfig;
                                }
                                var givenRuntimeConfig = JSON.parse(clientSupportedSDKVersion);
                                initializeHelperLogger('givenRuntimeConfig parsed to %o', givenRuntimeConfig !== null && givenRuntimeConfig !== void 0 ? givenRuntimeConfig : 'null');
                                if (!givenRuntimeConfig) {
                                    throw new Error('givenRuntimeConfig string was successfully parsed. However, it parsed to value of null');
                                }
                                else {
                                    applyRuntimeConfig(givenRuntimeConfig);
                                }
                            }
                            catch (e) {
                                if (e instanceof SyntaxError) {
                                    applyRuntimeConfig(generateBackCompatRuntimeConfig(GlobalVars.clientSupportedSDKVersion));
                                }
                                else {
                                    throw e;
                                }
                            }
                        }
                        else {
                            // If it's any error that's not a JSON parsing error, we want the program to fail.
                            throw e;
                        }
                    }
                    GlobalVars.initializeCompleted = true;
                });
                authentication.initialize();
                menus.initialize();
                pages.config.initialize();
                dialog.initialize();
            }
            // Handle additional valid message origins if specified
            if (Array.isArray(validMessageOrigins)) {
                processAdditionalValidOrigins(validMessageOrigins);
            }
            resolve(GlobalVars.initializePromise);
        });
    }
    /**
     * @hidden
     * Undocumented function used to set a mock window for unit tests
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function _initialize(hostWindow) {
        Communication.currentWindow = hostWindow;
    }
    app._initialize = _initialize;
    /**
     * @hidden
     * Undocumented function used to clear state between unit tests
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function _uninitialize() {
        if (!GlobalVars.initializeCalled) {
            return;
        }
        if (GlobalVars.frameContext) {
            /* eslint-disable strict-null-checks/all */ /* Fix tracked by 5730662 */
            registerOnThemeChangeHandler(null);
            pages.backStack.registerBackButtonHandler(null);
            pages.registerFullScreenHandler(null);
            teamsCore.registerBeforeUnloadHandler(null);
            teamsCore.registerOnLoadHandler(null);
            logs.registerGetLogHandler(null); /* Fix tracked by 5730662 */
            /* eslint-enable strict-null-checks/all */
        }
        if (GlobalVars.frameContext === FrameContexts.settings) {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            pages.config.registerOnSaveHandler(null);
        }
        if (GlobalVars.frameContext === FrameContexts.remove) {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            pages.config.registerOnRemoveHandler(null);
        }
        GlobalVars.initializeCalled = false;
        GlobalVars.initializeCompleted = false;
        GlobalVars.initializePromise = null;
        GlobalVars.additionalValidOrigins = [];
        GlobalVars.frameContext = null;
        GlobalVars.hostClientType = null;
        GlobalVars.isFramelessWindow = false;
        uninitializeCommunication();
    }
    app._uninitialize = _uninitialize;
    /**
     * Retrieves the current context the frame is running in.
     *
     * @returns Promise that will resolve with the {@link app.Context} object.
     */
    function getContext() {
        return new Promise(function (resolve) {
            ensureInitializeCalled();
            resolve(sendAndUnwrap('getContext'));
        }).then(function (legacyContext) { return transformLegacyContextToAppContext(legacyContext); }); // converts globalcontext to app.context
    }
    app.getContext = getContext;
    /**
     * Notifies the frame that app has loaded and to hide the loading indicator if one is shown.
     */
    function notifyAppLoaded() {
        ensureInitializeCalled();
        sendMessageToParent(app.Messages.AppLoaded, [version]);
    }
    app.notifyAppLoaded = notifyAppLoaded;
    /**
     * Notifies the frame that app initialization is successful and is ready for user interaction.
     */
    function notifySuccess() {
        ensureInitializeCalled();
        sendMessageToParent(app.Messages.Success, [version]);
    }
    app.notifySuccess = notifySuccess;
    /**
     * Notifies the frame that app initialization has failed and to show an error page in its place.
     *
     * @param appInitializationFailedRequest - The failure request containing the reason for why the app failed
     * during initialization as well as an optional message.
     */
    function notifyFailure(appInitializationFailedRequest) {
        ensureInitializeCalled();
        sendMessageToParent(app.Messages.Failure, [
            appInitializationFailedRequest.reason,
            appInitializationFailedRequest.message,
        ]);
    }
    app.notifyFailure = notifyFailure;
    /**
     * Notifies the frame that app initialized with some expected errors.
     *
     * @param expectedFailureRequest - The expected failure request containing the reason and an optional message
     */
    function notifyExpectedFailure(expectedFailureRequest) {
        ensureInitializeCalled();
        sendMessageToParent(app.Messages.ExpectedFailure, [expectedFailureRequest.reason, expectedFailureRequest.message]);
    }
    app.notifyExpectedFailure = notifyExpectedFailure;
    /**
     * Registers a handler for theme changes.
     *
     * @remarks
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the user changes their theme.
     */
    function registerOnThemeChangeHandler(handler) {
        // allow for registration cleanup even when not called initialize
        handler && ensureInitializeCalled();
        handlers_registerOnThemeChangeHandler(handler);
    }
    app.registerOnThemeChangeHandler = registerOnThemeChangeHandler;
    /**
     * open link API.
     *
     * @param deepLink - deep link.
     * @returns Promise that will be fulfilled when the operation has completed
     */
    function openLink(deepLink) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
            resolve(sendAndHandleStatusAndReason('executeDeepLink', deepLink));
        });
    }
    app.openLink = openLink;
})(app || (app = {}));
/**
 * @hidden
 * Transforms the Legacy Context object received from Messages to the structured app.Context object
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function transformLegacyContextToAppContext(legacyContext) {
    var context = {
        actionInfo: legacyContext.actionInfo,
        app: {
            locale: legacyContext.locale,
            sessionId: legacyContext.appSessionId ? legacyContext.appSessionId : '',
            theme: legacyContext.theme ? legacyContext.theme : 'default',
            iconPositionVertical: legacyContext.appIconPosition,
            osLocaleInfo: legacyContext.osLocaleInfo,
            parentMessageId: legacyContext.parentMessageId,
            userClickTime: legacyContext.userClickTime,
            userFileOpenPreference: legacyContext.userFileOpenPreference,
            host: {
                name: legacyContext.hostName ? legacyContext.hostName : HostName.teams,
                clientType: legacyContext.hostClientType ? legacyContext.hostClientType : HostClientType.web,
                sessionId: legacyContext.sessionId ? legacyContext.sessionId : '',
                ringId: legacyContext.ringId,
            },
            appLaunchId: legacyContext.appLaunchId,
        },
        page: {
            id: legacyContext.entityId,
            frameContext: legacyContext.frameContext ? legacyContext.frameContext : GlobalVars.frameContext,
            subPageId: legacyContext.subEntityId,
            isFullScreen: legacyContext.isFullScreen,
            isMultiWindow: legacyContext.isMultiWindow,
            sourceOrigin: legacyContext.sourceOrigin,
        },
        user: {
            id: legacyContext.userObjectId,
            displayName: legacyContext.userDisplayName,
            isCallingAllowed: legacyContext.isCallingAllowed,
            isPSTNCallingAllowed: legacyContext.isPSTNCallingAllowed,
            licenseType: legacyContext.userLicenseType,
            loginHint: legacyContext.loginHint,
            userPrincipalName: legacyContext.userPrincipalName,
            tenant: legacyContext.tid
                ? {
                    id: legacyContext.tid,
                    teamsSku: legacyContext.tenantSKU,
                }
                : undefined,
        },
        channel: legacyContext.channelId
            ? {
                id: legacyContext.channelId,
                displayName: legacyContext.channelName,
                relativeUrl: legacyContext.channelRelativeUrl,
                membershipType: legacyContext.channelType,
                defaultOneNoteSectionId: legacyContext.defaultOneNoteSectionId,
                ownerGroupId: legacyContext.hostTeamGroupId,
                ownerTenantId: legacyContext.hostTeamTenantId,
            }
            : undefined,
        chat: legacyContext.chatId
            ? {
                id: legacyContext.chatId,
            }
            : undefined,
        meeting: legacyContext.meetingId
            ? {
                id: legacyContext.meetingId,
            }
            : undefined,
        sharepoint: legacyContext.sharepoint,
        team: legacyContext.teamId
            ? {
                internalId: legacyContext.teamId,
                displayName: legacyContext.teamName,
                type: legacyContext.teamType,
                groupId: legacyContext.groupId,
                templateId: legacyContext.teamTemplateId,
                isArchived: legacyContext.isTeamArchived,
                userRole: legacyContext.userTeamRole,
            }
            : undefined,
        sharePointSite: legacyContext.teamSiteUrl ||
            legacyContext.teamSiteDomain ||
            legacyContext.teamSitePath ||
            legacyContext.mySitePath ||
            legacyContext.mySiteDomain
            ? {
                teamSiteUrl: legacyContext.teamSiteUrl,
                teamSiteDomain: legacyContext.teamSiteDomain,
                teamSitePath: legacyContext.teamSitePath,
                teamSiteId: legacyContext.teamSiteId,
                mySitePath: legacyContext.mySitePath,
                mySiteDomain: legacyContext.mySiteDomain,
            }
            : undefined,
    };
    return context;
}

;// CONCATENATED MODULE: ./src/public/pages.ts







/**
 * Navigation-specific part of the SDK.
 */
var pages;
(function (pages) {
    /**
     * Return focus to the host. Will move focus forward or backward based on where the application container falls in
     * the F6/tab order in the host.
     * On mobile hosts or hosts where there is no keyboard interaction or UI notion of "focus" this function has no
     * effect and will be a no-op when called.
     * @param navigateForward - Determines the direction to focus in host.
     */
    function returnFocus(navigateForward) {
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('returnFocus', [navigateForward]);
    }
    pages.returnFocus = returnFocus;
    /**
     * @hidden
     *
     * Registers a handler for specifying focus when it passes from the host to the application.
     * On mobile hosts or hosts where there is no UI notion of "focus" the handler registered with
     * this function will never be called.
     *
     * @param handler - The handler for placing focus within the application.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerFocusEnterHandler(handler) {
        registerHandlerHelper('focusEnter', handler, [], function () {
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
        });
    }
    pages.registerFocusEnterHandler = registerFocusEnterHandler;
    /**
     * Sets/Updates the current frame with new information
     *
     * @param frameInfo - Frame information containing the URL used in the iframe on reload and the URL for when the
     * user clicks 'Go To Website'
     */
    function setCurrentFrame(frameInfo) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('setFrameContext', [frameInfo]);
    }
    pages.setCurrentFrame = setCurrentFrame;
    /**
     * Initializes the library with context information for the frame
     *
     * @param frameInfo - Frame information containing the URL used in the iframe on reload and the URL for when the
     *  user clicks 'Go To Website'
     * @param callback - An optional callback that is executed once the application has finished initialization.
     * @param validMessageOrigins - An optional list of cross-frame message origins. They must have
     * https: protocol otherwise they will be ignored. Example: https:www.example.com
     */
    function initializeWithFrameContext(frameInfo, callback, validMessageOrigins) {
        app.initialize(validMessageOrigins).then(function () { return callback && callback(); });
        setCurrentFrame(frameInfo);
    }
    pages.initializeWithFrameContext = initializeWithFrameContext;
    /**
     * Gets the config for the current instance.
     * @returns Promise that resolves with the {@link InstanceConfig} object.
     */
    function getConfig() {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.settings, FrameContexts.remove, FrameContexts.sidePanel);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            resolve(sendAndUnwrap('settings.getSettings'));
        });
    }
    pages.getConfig = getConfig;
    /**
     * Navigates the frame to a new cross-domain URL. The domain of this URL must match at least one of the
     * valid domains specified in the validDomains block of the manifest; otherwise, an exception will be
     * thrown. This function needs to be used only when navigating the frame to a URL in a different domain
     * than the current one in a way that keeps the application informed of the change and allows the SDK to
     * continue working.
     * @param url - The URL to navigate the frame to.
     * @returns Promise that resolves when the navigation has completed.
     */
    function navigateCrossDomain(url) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.remove, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            var errorMessage = 'Cross-origin navigation is only supported for URLs matching the pattern registered in the manifest.';
            resolve(sendAndHandleStatusAndReasonWithDefaultError('navigateCrossDomain', errorMessage, url));
        });
    }
    pages.navigateCrossDomain = navigateCrossDomain;
    /**
     * Navigate to the given application ID and page ID, with optional parameters for a WebURL (if the application
     * cannot be navigated to, such as if it is not installed), Channel ID (for applications installed as a channel tab),
     * and sub-page ID (for navigating to specific content within the page). This is equivalent to navigating to
     * a deep link with the above data, but does not require the application to build a URL or worry about different
     * deep link formats for different hosts.
     * @param params - Parameters for the navigation
     * @returns a promise that will resolve if the navigation was successful
     */
    function navigateToApp(params) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (runtime.isLegacyTeams) {
                resolve(sendAndHandleStatusAndReason('executeDeepLink', createTeamsAppLink(params)));
            }
            else {
                resolve(sendAndHandleStatusAndReason('pages.navigateToApp', params));
            }
        });
    }
    pages.navigateToApp = navigateToApp;
    /**
     * Shares a deep link that a user can use to navigate back to a specific state in this page.
     * Please note that this method does yet work on mobile hosts.
     *
     * @param deepLinkParameters - ID and label for the link and fallback URL.
     */
    function shareDeepLink(deepLinkParameters) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('shareDeepLink', [
            deepLinkParameters.subPageId,
            deepLinkParameters.subPageLabel,
            deepLinkParameters.subPageWebUrl,
        ]);
    }
    pages.shareDeepLink = shareDeepLink;
    /**
     * Registers a handler for changes from or to full-screen view for a tab.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     * On hosts where there is no support for making an app full screen, the handler registered
     * with this function will never be called.
     * @param handler - The handler to invoke when the user toggles full-screen view for a tab.
     */
    function registerFullScreenHandler(handler) {
        registerHandlerHelper('fullScreenChange', handler, [], function () {
            if (handler && !isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
        });
    }
    pages.registerFullScreenHandler = registerFullScreenHandler;
    /**
     * Checks if the pages capability is supported by the host
     * @returns boolean to represent whether the appEntity capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.pages ? true : false;
    }
    pages.isSupported = isSupported;
    /**
     * Provides APIs for querying and navigating between contextual tabs of an application. Unlike personal tabs,
     * contextual tabs are pages associated with a specific context, such as channel or chat.
     */
    var tabs;
    (function (tabs) {
        /**
         * Navigates the hosted application to the specified tab instance.
         * @param tabInstance - The destination tab instance.
         * @returns Promise that resolves when the navigation has completed.
         */
        function navigateToTab(tabInstance) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                var errorMessage = 'Invalid internalTabInstanceId and/or channelId were/was provided';
                resolve(sendAndHandleStatusAndReasonWithDefaultError('navigateToTab', errorMessage, tabInstance));
            });
        }
        tabs.navigateToTab = navigateToTab;
        /**
         * Retrieves application tabs for the current user.
         * If no TabInstanceParameters are passed, the application defaults to favorite teams and favorite channels.
         * @param tabInstanceParameters - An optional set of flags that specify whether to scope call to favorite teams or channels.
         * @returns Promise that resolves with the {@link TabInformation}. Contains information for the user's tabs that are owned by this application {@link TabInstance}.
         */
        function getTabInstances(tabInstanceParameters) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                resolve(sendAndUnwrap('getTabInstances', tabInstanceParameters));
            });
        }
        tabs.getTabInstances = getTabInstances;
        /**
         * Retrieves the most recently used application tabs for the current user.
         * @param tabInstanceParameters - An optional set of flags. Note this is currently ignored and kept for future use.
         * @returns Promise that resolves with the {@link TabInformation}. Contains information for the users' most recently used tabs {@link TabInstance}.
         */
        function getMruTabInstances(tabInstanceParameters) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                resolve(sendAndUnwrap('getMruTabInstances', tabInstanceParameters));
            });
        }
        tabs.getMruTabInstances = getMruTabInstances;
        /**
         * Checks if the pages.tab capability is supported by the host
         * @returns boolean to represent whether the pages.tab capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.tabs
                    ? true
                    : false
                : false;
        }
        tabs.isSupported = isSupported;
    })(tabs = pages.tabs || (pages.tabs = {}));
    /**
     * Provides APIs to interact with the configuration-specific part of the SDK.
     * This object is usable only on the configuration frame.
     */
    var config;
    (function (config) {
        var saveHandler;
        var removeHandler;
        /**
         * @hidden
         * Hide from docs because this function is only used during initialization
         *
         * Adds register handlers for settings.save and settings.remove upon initialization. Function is called in {@link app.initializeHelper}
         * @internal
         * Limited to Microsoft-internal use
         */
        function initialize() {
            registerHandler('settings.save', handleSave, false);
            registerHandler('settings.remove', handleRemove, false);
        }
        config.initialize = initialize;
        /**
         * Sets the validity state for the configuration.
         * The initial value is false, so the user cannot save the configuration until this is called with true.
         * @param validityState - Indicates whether the save or remove button is enabled for the user.
         */
        function setValidityState(validityState) {
            ensureInitialized(runtime, FrameContexts.settings, FrameContexts.remove);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('settings.setValidityState', [validityState]);
        }
        config.setValidityState = setValidityState;
        /**
         * Sets the configuration for the current instance.
         * This is an asynchronous operation; calls to getConfig are not guaranteed to reflect the changed state.
         * @param instanceConfig - The desired configuration for this instance.
         * @returns Promise that resolves when the operation has completed.
         */
        function setConfig(instanceConfig) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.settings, FrameContexts.sidePanel);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                resolve(sendAndHandleStatusAndReason('settings.setSettings', instanceConfig));
            });
        }
        config.setConfig = setConfig;
        /**
         * Registers a handler for when the user attempts to save the configuration. This handler should be used
         * to create or update the underlying resource powering the content.
         * The object passed to the handler must be used to notify whether to proceed with the save.
         * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
         * @param handler - The handler to invoke when the user selects the Save button.
         */
        function registerOnSaveHandler(handler) {
            registerOnSaveHandlerHelper(handler, function () {
                if (handler && !isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        config.registerOnSaveHandler = registerOnSaveHandler;
        /**
         * @hidden
         * Undocumented helper function with shared code between deprecated version and current version of the registerOnSaveHandler API.
         *
         * @internal
         * Limited to Microsoft-internal use
         *
         * @param handler - The handler to invoke when the user selects the Save button.
         * @param versionSpecificHelper - The helper function containing logic pertaining to a specific version of the API.
         */
        function registerOnSaveHandlerHelper(handler, versionSpecificHelper) {
            // allow for registration cleanup even when not finished initializing
            handler && ensureInitialized(runtime, FrameContexts.settings);
            if (versionSpecificHelper) {
                versionSpecificHelper();
            }
            saveHandler = handler;
            handler && sendMessageToParent('registerHandler', ['save']);
        }
        config.registerOnSaveHandlerHelper = registerOnSaveHandlerHelper;
        /**
         * Registers a handler for user attempts to remove content. This handler should be used
         * to remove the underlying resource powering the content.
         * The object passed to the handler must be used to indicate whether to proceed with the removal.
         * Only one handler may be registered at a time. Subsequent registrations will override the first.
         * @param handler - The handler to invoke when the user selects the Remove button.
         */
        function registerOnRemoveHandler(handler) {
            registerOnRemoveHandlerHelper(handler, function () {
                if (handler && !isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        config.registerOnRemoveHandler = registerOnRemoveHandler;
        /**
         * @hidden
         * Undocumented helper function with shared code between deprecated version and current version of the registerOnRemoveHandler API.
         *
         * @internal
         * Limited to Microsoft-internal use
         *
         * @param handler - The handler to invoke when the user selects the Remove button.
         * @param versionSpecificHelper - The helper function containing logic pertaining to a specific version of the API.
         */
        function registerOnRemoveHandlerHelper(handler, versionSpecificHelper) {
            // allow for registration cleanup even when not finished initializing
            handler && ensureInitialized(runtime, FrameContexts.remove, FrameContexts.settings);
            if (versionSpecificHelper) {
                versionSpecificHelper();
            }
            removeHandler = handler;
            handler && sendMessageToParent('registerHandler', ['remove']);
        }
        config.registerOnRemoveHandlerHelper = registerOnRemoveHandlerHelper;
        function handleSave(result) {
            var saveEventType = new SaveEventImpl(result);
            if (saveHandler) {
                saveHandler(saveEventType);
            }
            else if (Communication.childWindow) {
                sendMessageEventToChild('settings.save', [result]);
            }
            else {
                // If no handler is registered, we assume success.
                saveEventType.notifySuccess();
            }
        }
        /**
         * Registers a handler for when the tab configuration is changed by the user
         * @param handler - The handler to invoke when the user clicks on Settings.
         */
        function registerChangeConfigHandler(handler) {
            registerHandlerHelper('changeSettings', handler, [FrameContexts.content], function () {
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        config.registerChangeConfigHandler = registerChangeConfigHandler;
        /**
         * @hidden
         * Hide from docs, since this class is not directly used.
         */
        var SaveEventImpl = /** @class */ (function () {
            function SaveEventImpl(result) {
                this.notified = false;
                this.result = result ? result : {};
            }
            SaveEventImpl.prototype.notifySuccess = function () {
                this.ensureNotNotified();
                sendMessageToParent('settings.save.success');
                this.notified = true;
            };
            SaveEventImpl.prototype.notifyFailure = function (reason) {
                this.ensureNotNotified();
                sendMessageToParent('settings.save.failure', [reason]);
                this.notified = true;
            };
            SaveEventImpl.prototype.ensureNotNotified = function () {
                if (this.notified) {
                    throw new Error('The SaveEvent may only notify success or failure once.');
                }
            };
            return SaveEventImpl;
        }());
        function handleRemove() {
            var removeEventType = new RemoveEventImpl();
            if (removeHandler) {
                removeHandler(removeEventType);
            }
            else if (Communication.childWindow) {
                sendMessageEventToChild('settings.remove', []);
            }
            else {
                // If no handler is registered, we assume success.
                removeEventType.notifySuccess();
            }
        }
        /**
         * @hidden
         * Hide from docs, since this class is not directly used.
         */
        var RemoveEventImpl = /** @class */ (function () {
            function RemoveEventImpl() {
                this.notified = false;
            }
            RemoveEventImpl.prototype.notifySuccess = function () {
                this.ensureNotNotified();
                sendMessageToParent('settings.remove.success');
                this.notified = true;
            };
            RemoveEventImpl.prototype.notifyFailure = function (reason) {
                this.ensureNotNotified();
                sendMessageToParent('settings.remove.failure', [reason]);
                this.notified = true;
            };
            RemoveEventImpl.prototype.ensureNotNotified = function () {
                if (this.notified) {
                    throw new Error('The removeEventType may only notify success or failure once.');
                }
            };
            return RemoveEventImpl;
        }());
        /**
         * Checks if the pages.config capability is supported by the host
         * @returns boolean to represent whether the pages.config capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.config
                    ? true
                    : false
                : false;
        }
        config.isSupported = isSupported;
    })(config = pages.config || (pages.config = {}));
    /**
     * Provides APIs for handling the user's navigational history.
     */
    var backStack;
    (function (backStack) {
        var backButtonPressHandler;
        /**
         * @hidden
         * Register backButtonPress handler.
         *
         * @internal
         * Limited to Microsoft-internal use.
         */
        function _initialize() {
            registerHandler('backButtonPress', handleBackButtonPress, false);
        }
        backStack._initialize = _initialize;
        /**
         * Navigates back in the hosted application. See {@link pages.backStack.registerBackButtonHandler} for notes on usage.
         * @returns Promise that resolves when the navigation has completed.
         */
        function navigateBack() {
            return new Promise(function (resolve) {
                ensureInitialized(runtime);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                var errorMessage = 'Back navigation is not supported in the current client or context.';
                resolve(sendAndHandleStatusAndReasonWithDefaultError('navigateBack', errorMessage));
            });
        }
        backStack.navigateBack = navigateBack;
        /**
         * Registers a handler for user presses of the host client's back button. Experiences that maintain an internal
         * navigation stack should use this handler to navigate the user back within their frame. If an application finds
         * that after running its back button handler it cannot handle the event it should call the navigateBack
         * method to ask the host client to handle it instead.
         * @param handler - The handler to invoke when the user presses the host client's back button.
         */
        function registerBackButtonHandler(handler) {
            registerBackButtonHandlerHelper(handler, function () {
                if (handler && !isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        backStack.registerBackButtonHandler = registerBackButtonHandler;
        /**
         * @hidden
         * Undocumented helper function with shared code between deprecated version and current version of the registerBackButtonHandler API.
         *
         * @internal
         * Limited to Microsoft-internal use
         *
         * @param handler - The handler to invoke when the user presses the host client's back button.
         * @param versionSpecificHelper - The helper function containing logic pertaining to a specific version of the API.
         */
        function registerBackButtonHandlerHelper(handler, versionSpecificHelper) {
            // allow for registration cleanup even when not finished initializing
            handler && ensureInitialized(runtime);
            if (versionSpecificHelper) {
                versionSpecificHelper();
            }
            backButtonPressHandler = handler;
            handler && sendMessageToParent('registerHandler', ['backButton']);
        }
        backStack.registerBackButtonHandlerHelper = registerBackButtonHandlerHelper;
        function handleBackButtonPress() {
            if (!backButtonPressHandler || !backButtonPressHandler()) {
                if (Communication.childWindow) {
                    // If the current window did not handle it let the child window
                    sendMessageEventToChild('backButtonPress', []);
                }
                else {
                    navigateBack();
                }
            }
        }
        /**
         * Checks if the pages.backStack capability is supported by the host
         * @returns boolean to represent whether the pages.backStack capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.backStack
                    ? true
                    : false
                : false;
        }
        backStack.isSupported = isSupported;
    })(backStack = pages.backStack || (pages.backStack = {}));
    /**
     * @hidden
     * Hide from docs
     * ------
     * Provides APIs to interact with the full-trust part of the SDK. Limited to 1P applications
     */
    var fullTrust;
    (function (fullTrust) {
        /**
         * @hidden
         * Hide from docs
         * ------
         * Place the tab into full-screen mode.
         */
        function enterFullscreen() {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('enterFullscreen', []);
        }
        fullTrust.enterFullscreen = enterFullscreen;
        /**
         * @hidden
         * Hide from docs
         * ------
         * Reverts the tab into normal-screen mode.
         */
        function exitFullscreen() {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            sendMessageToParent('exitFullscreen', []);
        }
        fullTrust.exitFullscreen = exitFullscreen;
        /**
         * @hidden
         *
         * Checks if the pages.fullTrust capability is supported by the host
         * @returns boolean to represent whether the pages.fullTrust capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.fullTrust
                    ? true
                    : false
                : false;
        }
        fullTrust.isSupported = isSupported;
    })(fullTrust = pages.fullTrust || (pages.fullTrust = {}));
    /**
     * Provides APIs to interact with the app button part of the SDK.
     */
    var appButton;
    (function (appButton) {
        /**
         * Registers a handler for clicking the app button.
         * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
         * @param handler - The handler to invoke when the personal app button is clicked in the app bar.
         */
        function onClick(handler) {
            registerHandlerHelper('appButtonClick', handler, [FrameContexts.content], function () {
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        appButton.onClick = onClick;
        /**
         * Registers a handler for entering hover of the app button.
         * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
         * @param handler - The handler to invoke when entering hover of the personal app button in the app bar.
         */
        function onHoverEnter(handler) {
            registerHandlerHelper('appButtonHoverEnter', handler, [FrameContexts.content], function () {
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        appButton.onHoverEnter = onHoverEnter;
        /**
         * Registers a handler for exiting hover of the app button.
         * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
         * @param handler - The handler to invoke when exiting hover of the personal app button in the app bar.
         */
        function onHoverLeave(handler) {
            registerHandlerHelper('appButtonHoverLeave', handler, [FrameContexts.content], function () {
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
            });
        }
        appButton.onHoverLeave = onHoverLeave;
        /**
         * Checks if pages.appButton capability is supported by the host
         * @returns boolean to represent whether the pages.appButton capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.appButton
                    ? true
                    : false
                : false;
        }
        appButton.isSupported = isSupported;
    })(appButton = pages.appButton || (pages.appButton = {}));
    /**
     * Provides functions for navigating without needing to specify your application ID.
     *
     * @beta
     */
    var currentApp;
    (function (currentApp) {
        /**
         * Navigate within the currently running application with page ID, and sub-page ID (for navigating to
         * specific content within the page).
         * @param params - Parameters for the navigation
         * @returns a promise that will resolve if the navigation was successful
         *
         * @beta
         */
        function navigateTo(params) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                resolve(sendAndHandleSdkError('pages.currentApp.navigateTo', params));
            });
        }
        currentApp.navigateTo = navigateTo;
        /**
         * Navigate to the currently running application's first static page defined in the application
         * manifest.
         * @beta
         */
        function navigateToDefaultPage() {
            return new Promise(function (resolve) {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                resolve(sendAndHandleSdkError('pages.currentApp.navigateToDefaultPage'));
            });
        }
        currentApp.navigateToDefaultPage = navigateToDefaultPage;
        /**
         * Checks if pages.currentApp capability is supported by the host
         * @returns boolean to represent whether the pages.currentApp capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @beta
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.pages
                ? runtime.supports.pages.currentApp
                    ? true
                    : false
                : false;
        }
        currentApp.isSupported = isSupported;
    })(currentApp = pages.currentApp || (pages.currentApp = {}));
})(pages || (pages = {}));

;// CONCATENATED MODULE: ./src/internal/handlers.ts
/* eslint-disable @typescript-eslint/ban-types */
var __spreadArray = (undefined && undefined.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};





var handlersLogger = getLogger('handlers');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
var HandlersPrivate = /** @class */ (function () {
    function HandlersPrivate() {
    }
    HandlersPrivate.handlers = {};
    return HandlersPrivate;
}());
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function initializeHandlers() {
    // ::::::::::::::::::::MicrosoftTeams SDK Internal :::::::::::::::::
    HandlersPrivate.handlers['themeChange'] = handleThemeChange;
    HandlersPrivate.handlers['load'] = handleLoad;
    HandlersPrivate.handlers['beforeUnload'] = handleBeforeUnload;
    pages.backStack._initialize();
}
var callHandlerLogger = handlersLogger.extend('callHandler');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function callHandler(name, args) {
    var handler = HandlersPrivate.handlers[name];
    if (handler) {
        callHandlerLogger('Invoking the registered handler for message %s with arguments %o', name, args);
        var result = handler.apply(this, args);
        return [true, result];
    }
    else if (Communication.childWindow) {
        sendMessageEventToChild(name, args);
        return [false, undefined];
    }
    else {
        callHandlerLogger('Handler for action message %s not found.', name);
        return [false, undefined];
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function registerHandler(name, handler, sendMessage, args) {
    if (sendMessage === void 0) { sendMessage = true; }
    if (args === void 0) { args = []; }
    if (handler) {
        HandlersPrivate.handlers[name] = handler;
        sendMessage && sendMessageToParent('registerHandler', __spreadArray([name], args, true));
    }
    else {
        delete HandlersPrivate.handlers[name];
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function removeHandler(name) {
    delete HandlersPrivate.handlers[name];
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function doesHandlerExist(name) {
    return HandlersPrivate.handlers[name] != null;
}
/**
 * @hidden
 * Undocumented helper function with shared code between deprecated version and current version of register*Handler APIs
 *
 * @internal
 * Limited to Microsoft-internal use
 *
 * @param name - The name of the handler to register.
 * @param handler - The handler to invoke.
 * @param contexts - The context within which it is valid to register this handler.
 * @param registrationHelper - The helper function containing logic pertaining to a specific version of the API.
 */
function registerHandlerHelper(name, handler, contexts, registrationHelper) {
    // allow for registration cleanup even when not finished initializing
    handler && ensureInitialized.apply(void 0, __spreadArray([runtime], contexts, false));
    if (registrationHelper) {
        registrationHelper();
    }
    registerHandler(name, handler);
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handlers_registerOnThemeChangeHandler(handler) {
    HandlersPrivate.themeChangeHandler = handler;
    handler && sendMessageToParent('registerHandler', ['themeChange']);
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handleThemeChange(theme) {
    if (HandlersPrivate.themeChangeHandler) {
        HandlersPrivate.themeChangeHandler(theme);
    }
    if (Communication.childWindow) {
        sendMessageEventToChild('themeChange', [theme]);
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handlers_registerOnLoadHandler(handler) {
    HandlersPrivate.loadHandler = handler;
    handler && sendMessageToParent('registerHandler', ['load']);
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handleLoad(context) {
    if (HandlersPrivate.loadHandler) {
        HandlersPrivate.loadHandler(context);
    }
    if (Communication.childWindow) {
        sendMessageEventToChild('load', [context]);
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handlers_registerBeforeUnloadHandler(handler) {
    HandlersPrivate.beforeUnloadHandler = handler;
    handler && sendMessageToParent('registerHandler', ['beforeUnload']);
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handleBeforeUnload() {
    var readyToUnload = function () {
        sendMessageToParent('readyToUnload', []);
    };
    if (!HandlersPrivate.beforeUnloadHandler || !HandlersPrivate.beforeUnloadHandler(readyToUnload)) {
        if (Communication.childWindow) {
            sendMessageEventToChild('beforeUnload');
        }
        else {
            readyToUnload();
        }
    }
}

;// CONCATENATED MODULE: ./src/internal/communication.ts
/* eslint-disable @typescript-eslint/ban-types */
/* eslint-disable @typescript-eslint/no-explicit-any */
var communication_spreadArray = (undefined && undefined.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};






var communicationLogger = getLogger('communication');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
var Communication = /** @class */ (function () {
    function Communication() {
    }
    return Communication;
}());

/**
 * @internal
 * Limited to Microsoft-internal use
 */
var CommunicationPrivate = /** @class */ (function () {
    function CommunicationPrivate() {
    }
    CommunicationPrivate.parentMessageQueue = [];
    CommunicationPrivate.childMessageQueue = [];
    CommunicationPrivate.nextMessageId = 0;
    CommunicationPrivate.callbacks = {};
    CommunicationPrivate.promiseCallbacks = {};
    return CommunicationPrivate;
}());
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function initializeCommunication(validMessageOrigins) {
    // Listen for messages post to our window
    CommunicationPrivate.messageListener = function (evt) { return processMessage(evt); };
    // If we are in an iframe, our parent window is the one hosting us (i.e., window.parent); otherwise,
    // it's the window that opened us (i.e., window.opener)
    Communication.currentWindow = Communication.currentWindow || window;
    Communication.parentWindow =
        Communication.currentWindow.parent !== Communication.currentWindow.self
            ? Communication.currentWindow.parent
            : Communication.currentWindow.opener;
    // Listen to messages from the parent or child frame.
    // Frameless windows will only receive this event from child frames and if validMessageOrigins is passed.
    if (Communication.parentWindow || validMessageOrigins) {
        Communication.currentWindow.addEventListener('message', CommunicationPrivate.messageListener, false);
    }
    if (!Communication.parentWindow) {
        var extendedWindow = Communication.currentWindow;
        if (extendedWindow.nativeInterface) {
            GlobalVars.isFramelessWindow = true;
            extendedWindow.onNativeMessage = handleParentMessage;
        }
        else {
            // at this point we weren't able to find a parent to talk to, no way initialization will succeed
            return Promise.reject(new Error('Initialization Failed. No Parent window found.'));
        }
    }
    try {
        // Send the initialized message to any origin, because at this point we most likely don't know the origin
        // of the parent window, and this message contains no data that could pose a security risk.
        Communication.parentOrigin = '*';
        return sendMessageToParentAsync('initialize', [
            version,
            latestRuntimeApiVersion,
        ]).then(function (_a) {
            var context = _a[0], clientType = _a[1], runtimeConfig = _a[2], clientSupportedSDKVersion = _a[3];
            return { context: context, clientType: clientType, runtimeConfig: runtimeConfig, clientSupportedSDKVersion: clientSupportedSDKVersion };
        });
    }
    finally {
        Communication.parentOrigin = null;
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function uninitializeCommunication() {
    if (Communication.currentWindow) {
        Communication.currentWindow.removeEventListener('message', CommunicationPrivate.messageListener, false);
    }
    Communication.currentWindow = null;
    Communication.parentWindow = null;
    Communication.parentOrigin = null;
    Communication.childWindow = null;
    Communication.childOrigin = null;
    CommunicationPrivate.parentMessageQueue = [];
    CommunicationPrivate.childMessageQueue = [];
    CommunicationPrivate.nextMessageId = 0;
    CommunicationPrivate.callbacks = {};
    CommunicationPrivate.promiseCallbacks = {};
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function sendAndUnwrap(actionName) {
    var args = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        args[_i - 1] = arguments[_i];
    }
    return sendMessageToParentAsync(actionName, args).then(function (_a) {
        var result = _a[0];
        return result;
    });
}
function sendAndHandleStatusAndReason(actionName) {
    var args = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        args[_i - 1] = arguments[_i];
    }
    return sendMessageToParentAsync(actionName, args).then(function (_a) {
        var wasSuccessful = _a[0], reason = _a[1];
        if (!wasSuccessful) {
            throw new Error(reason);
        }
    });
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function sendAndHandleStatusAndReasonWithDefaultError(actionName, defaultError) {
    var args = [];
    for (var _i = 2; _i < arguments.length; _i++) {
        args[_i - 2] = arguments[_i];
    }
    return sendMessageToParentAsync(actionName, args).then(function (_a) {
        var wasSuccessful = _a[0], reason = _a[1];
        if (!wasSuccessful) {
            throw new Error(reason ? reason : defaultError);
        }
    });
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function sendAndHandleSdkError(actionName) {
    var args = [];
    for (var _i = 1; _i < arguments.length; _i++) {
        args[_i - 1] = arguments[_i];
    }
    return sendMessageToParentAsync(actionName, args).then(function (_a) {
        var error = _a[0], result = _a[1];
        if (error) {
            throw error;
        }
        return result;
    });
}
/**
 * @hidden
 * Send a message to parent. Uses nativeInterface on mobile to communicate with parent context
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function sendMessageToParentAsync(actionName, args) {
    if (args === void 0) { args = undefined; }
    return new Promise(function (resolve) {
        var request = sendMessageToParentHelper(actionName, args);
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        resolve(waitForResponse(request.id));
    });
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function waitForResponse(requestId) {
    return new Promise(function (resolve) {
        CommunicationPrivate.promiseCallbacks[requestId] = resolve;
    });
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function sendMessageToParent(actionName, argsOrCallback, callback) {
    var args;
    if (argsOrCallback instanceof Function) {
        callback = argsOrCallback;
    }
    else if (argsOrCallback instanceof Array) {
        args = argsOrCallback;
    }
    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
    var request = sendMessageToParentHelper(actionName, args);
    if (callback) {
        CommunicationPrivate.callbacks[request.id] = callback;
    }
}
var sendMessageToParentHelperLogger = communicationLogger.extend('sendMessageToParentHelper');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function sendMessageToParentHelper(actionName, args) {
    var logger = sendMessageToParentHelperLogger;
    var targetWindow = Communication.parentWindow;
    var request = createMessageRequest(actionName, args);
    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
    logger('Message %i information: %o', request.id, { actionName: actionName, args: args });
    if (GlobalVars.isFramelessWindow) {
        if (Communication.currentWindow && Communication.currentWindow.nativeInterface) {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            logger('Sending message %i to parent via framelessPostMessage interface', request.id);
            Communication.currentWindow.nativeInterface.framelessPostMessage(JSON.stringify(request));
        }
    }
    else {
        var targetOrigin = getTargetOrigin(targetWindow);
        // If the target window isn't closed and we already know its origin, send the message right away; otherwise,
        // queue the message and send it after the origin is established
        if (targetWindow && targetOrigin) {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            logger('Sending message %i to parent via postMessage', request.id);
            targetWindow.postMessage(request, targetOrigin);
        }
        else {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            logger('Adding message %i to parent message queue', request.id);
            getTargetMessageQueue(targetWindow).push(request);
        }
    }
    return request;
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function processMessage(evt) {
    // Process only if we received a valid message
    if (!evt || !evt.data || typeof evt.data !== 'object') {
        return;
    }
    // Process only if the message is coming from a different window and a valid origin
    // valid origins are either a pre-known
    var messageSource = evt.source || (evt.originalEvent && evt.originalEvent.source);
    var messageOrigin = evt.origin || (evt.originalEvent && evt.originalEvent.origin);
    if (!shouldProcessMessage(messageSource, messageOrigin)) {
        return;
    }
    // Update our parent and child relationships based on this message
    updateRelationships(messageSource, messageOrigin);
    // Handle the message
    if (messageSource === Communication.parentWindow) {
        handleParentMessage(evt);
    }
    else if (messageSource === Communication.childWindow) {
        handleChildMessage(evt);
    }
}
/**
 * @hidden
 * Validates the message source and origin, if it should be processed
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function shouldProcessMessage(messageSource, messageOrigin) {
    // Process if message source is a different window and if origin is either in
    // Teams' pre-known whitelist or supplied as valid origin by user during initialization
    if (Communication.currentWindow && messageSource === Communication.currentWindow) {
        return false;
    }
    else if (Communication.currentWindow &&
        Communication.currentWindow.location &&
        messageOrigin &&
        messageOrigin === Communication.currentWindow.location.origin) {
        return true;
    }
    else {
        return validateOrigin(new URL(messageOrigin));
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function updateRelationships(messageSource, messageOrigin) {
    // Determine whether the source of the message is our parent or child and update our
    // window and origin pointer accordingly
    // For frameless windows (i.e mobile), there is no parent frame, so the message must be from the child.
    if (!GlobalVars.isFramelessWindow &&
        (!Communication.parentWindow || Communication.parentWindow.closed || messageSource === Communication.parentWindow)) {
        Communication.parentWindow = messageSource;
        Communication.parentOrigin = messageOrigin;
    }
    else if (!Communication.childWindow ||
        Communication.childWindow.closed ||
        messageSource === Communication.childWindow) {
        Communication.childWindow = messageSource;
        Communication.childOrigin = messageOrigin;
    }
    // Clean up pointers to closed parent and child windows
    if (Communication.parentWindow && Communication.parentWindow.closed) {
        Communication.parentWindow = null;
        Communication.parentOrigin = null;
    }
    if (Communication.childWindow && Communication.childWindow.closed) {
        Communication.childWindow = null;
        Communication.childOrigin = null;
    }
    // If we have any messages in our queue, send them now
    flushMessageQueue(Communication.parentWindow);
    flushMessageQueue(Communication.childWindow);
}
var handleParentMessageLogger = communicationLogger.extend('handleParentMessage');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handleParentMessage(evt) {
    var logger = handleParentMessageLogger;
    if ('id' in evt.data && typeof evt.data.id === 'number') {
        // Call any associated Communication.callbacks
        var message = evt.data;
        var callback = CommunicationPrivate.callbacks[message.id];
        logger('Received a response from parent for message %i', message.id);
        if (callback) {
            logger('Invoking the registered callback for message %i with arguments %o', message.id, message.args);
            callback.apply(null, communication_spreadArray(communication_spreadArray([], message.args, true), [message.isPartialResponse], false));
            // Remove the callback to ensure that the callback is called only once and to free up memory if response is a complete response
            if (!isPartialResponse(evt)) {
                logger('Removing registered callback for message %i', message.id);
                delete CommunicationPrivate.callbacks[message.id];
            }
        }
        var promiseCallback = CommunicationPrivate.promiseCallbacks[message.id];
        if (promiseCallback) {
            logger('Invoking the registered promise callback for message %i with arguments %o', message.id, message.args);
            promiseCallback(message.args);
            logger('Removing registered promise callback for message %i', message.id);
            delete CommunicationPrivate.promiseCallbacks[message.id];
        }
    }
    else if ('func' in evt.data && typeof evt.data.func === 'string') {
        // Delegate the request to the proper handler
        var message = evt.data;
        logger('Received an action message %s from parent', message.func);
        callHandler(message.func, message.args);
    }
    else {
        logger('Received an unknown message: %O', evt);
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function isPartialResponse(evt) {
    return evt.data.isPartialResponse === true;
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function handleChildMessage(evt) {
    if ('id' in evt.data && 'func' in evt.data) {
        // Try to delegate the request to the proper handler, if defined
        var message_1 = evt.data;
        var _a = callHandler(message_1.func, message_1.args), called = _a[0], result = _a[1];
        if (called && typeof result !== 'undefined') {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            sendMessageResponseToChild(message_1.id, Array.isArray(result) ? result : [result]);
        }
        else {
            // No handler, proxy to parent
            sendMessageToParent(message_1.func, message_1.args, function () {
                var args = [];
                for (var _i = 0; _i < arguments.length; _i++) {
                    args[_i] = arguments[_i];
                }
                if (Communication.childWindow) {
                    var isPartialResponse_1 = args.pop();
                    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                    sendMessageResponseToChild(message_1.id, args, isPartialResponse_1);
                }
            });
        }
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function getTargetMessageQueue(targetWindow) {
    return targetWindow === Communication.parentWindow
        ? CommunicationPrivate.parentMessageQueue
        : targetWindow === Communication.childWindow
            ? CommunicationPrivate.childMessageQueue
            : [];
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function getTargetOrigin(targetWindow) {
    return targetWindow === Communication.parentWindow
        ? Communication.parentOrigin
        : targetWindow === Communication.childWindow
            ? Communication.childOrigin
            : null;
}
var flushMessageQueueLogger = communicationLogger.extend('flushMessageQueue');
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function flushMessageQueue(targetWindow) {
    var targetOrigin = getTargetOrigin(targetWindow);
    var targetMessageQueue = getTargetMessageQueue(targetWindow);
    var target = targetWindow == Communication.parentWindow ? 'parent' : 'child';
    while (targetWindow && targetOrigin && targetMessageQueue.length > 0) {
        var request = targetMessageQueue.shift();
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        flushMessageQueueLogger('Flushing message %i from ' + target + ' message queue via postMessage.', request.id);
        targetWindow.postMessage(request, targetOrigin);
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function waitForMessageQueue(targetWindow, callback) {
    var messageQueueMonitor = Communication.currentWindow.setInterval(function () {
        if (getTargetMessageQueue(targetWindow).length === 0) {
            clearInterval(messageQueueMonitor);
            callback();
        }
    }, 100);
}
/**
 * @hidden
 * Send a response to child for a message request that was from child
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function sendMessageResponseToChild(id, args, isPartialResponse) {
    var targetWindow = Communication.childWindow;
    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
    var response = createMessageResponse(id, args, isPartialResponse);
    var targetOrigin = getTargetOrigin(targetWindow);
    if (targetWindow && targetOrigin) {
        targetWindow.postMessage(response, targetOrigin);
    }
}
/**
 * @hidden
 * Send a custom message object that can be sent to child window,
 * instead of a response message to a child
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function sendMessageEventToChild(actionName, args) {
    var targetWindow = Communication.childWindow;
    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
    var customEvent = createMessageEvent(actionName, args);
    var targetOrigin = getTargetOrigin(targetWindow);
    // If the target window isn't closed and we already know its origin, send the message right away; otherwise,
    // queue the message and send it after the origin is established
    if (targetWindow && targetOrigin) {
        targetWindow.postMessage(customEvent, targetOrigin);
    }
    else {
        getTargetMessageQueue(targetWindow).push(customEvent);
    }
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function createMessageRequest(func, args) {
    return {
        id: CommunicationPrivate.nextMessageId++,
        func: func,
        timestamp: Date.now(),
        args: args || [],
    };
}
/**
 * @internal
 * Limited to Microsoft-internal use
 */
function createMessageResponse(id, args, isPartialResponse) {
    return {
        id: id,
        args: args || [],
        isPartialResponse: isPartialResponse,
    };
}
/**
 * @hidden
 * Creates a message object without any id, used for custom actions being sent to child frame/window
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function createMessageEvent(func, args) {
    return {
        func: func,
        args: args || [],
    };
}

;// CONCATENATED MODULE: ./src/private/logs.ts





/**
 * @hidden
 * Namespace to interact with the logging part of the SDK.
 * This object is used to send the app logs on demand to the host client
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var logs;
(function (logs) {
    /**
     * @hidden
     *
     * Registers a handler for getting app log
     *
     * @param handler - The handler to invoke to get the app log
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerGetLogHandler(handler) {
        // allow for registration cleanup even when not finished initializing
        handler && ensureInitialized(runtime);
        if (handler && !isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (handler) {
            registerHandler('log.request', function () {
                var log = handler();
                sendMessageToParent('log.receive', [log]);
            });
        }
        else {
            removeHandler('log.request');
        }
    }
    logs.registerGetLogHandler = registerGetLogHandler;
    /**
     * @hidden
     *
     * Checks if the logs capability is supported by the host
     * @returns boolean to represent whether the logs capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.logs ? true : false;
    }
    logs.isSupported = isSupported;
})(logs || (logs = {}));

;// CONCATENATED MODULE: ./src/private/interfaces.ts
/**
 * @hidden
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var NotificationTypes;
(function (NotificationTypes) {
    NotificationTypes["fileDownloadStart"] = "fileDownloadStart";
    NotificationTypes["fileDownloadComplete"] = "fileDownloadComplete";
})(NotificationTypes || (NotificationTypes = {}));
/**
 * @hidden
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var ViewerActionTypes;
(function (ViewerActionTypes) {
    ViewerActionTypes["view"] = "view";
    ViewerActionTypes["edit"] = "edit";
    ViewerActionTypes["editNew"] = "editNew";
})(ViewerActionTypes || (ViewerActionTypes = {}));
/**
 * @hidden
 *
 * User setting changes that can be subscribed to
 */
var UserSettingTypes;
(function (UserSettingTypes) {
    /**
     * @hidden
     * Use this key to subscribe to changes in user's file open preference
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    UserSettingTypes["fileOpenPreference"] = "fileOpenPreference";
    /**
     * @hidden
     * Use this key to subscribe to theme changes
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    UserSettingTypes["theme"] = "theme";
})(UserSettingTypes || (UserSettingTypes = {}));

;// CONCATENATED MODULE: ./src/private/privateAPIs.ts
/* eslint-disable @typescript-eslint/no-explicit-any */






/**
 * @hidden
 * Upload a custom App manifest directly to both team and personal scopes.
 * This method works just for the first party Apps.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function uploadCustomApp(manifestBlob, onComplete) {
    ensureInitialized(runtime);
    sendMessageToParent('uploadCustomApp', [manifestBlob], onComplete ? onComplete : getGenericOnCompleteHandler());
}
/**
 * @hidden
 * Sends a custom action MessageRequest to host or parent window
 *
 * @param actionName - Specifies name of the custom action to be sent
 * @param args - Specifies additional arguments passed to the action
 * @param callback - Optionally specify a callback to receive response parameters from the parent
 * @returns id of sent message
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function sendCustomMessage(actionName, args, callback) {
    ensureInitialized(runtime);
    sendMessageToParent(actionName, args, callback);
}
/**
 * @hidden
 * Sends a custom action MessageEvent to a child iframe/window, only if you are not using auth popup.
 * Otherwise it will go to the auth popup (which becomes the child)
 *
 * @param actionName - Specifies name of the custom action to be sent
 * @param args - Specifies additional arguments passed to the action
 * @returns id of sent message
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function sendCustomEvent(actionName, args) {
    ensureInitialized(runtime);
    //validate childWindow
    if (!Communication.childWindow) {
        throw new Error('The child window has not yet been initialized or is not present');
    }
    sendMessageEventToChild(actionName, args);
}
/**
 * @hidden
 * Adds a handler for an action sent by a child window or parent window
 *
 * @param actionName - Specifies name of the action message to handle
 * @param customHandler - The callback to invoke when the action message is received. The return value is sent to the child
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function registerCustomHandler(actionName, customHandler) {
    var _this = this;
    ensureInitialized(runtime);
    registerHandler(actionName, function () {
        var args = [];
        for (var _i = 0; _i < arguments.length; _i++) {
            args[_i] = arguments[_i];
        }
        return customHandler.apply(_this, args);
    });
}
/**
 * @hidden
 * register a handler to be called when a user setting changes. The changed setting type & value is provided in the callback.
 *
 * @param settingTypes - List of user setting changes to subscribe
 * @param handler - When a subscribed setting is updated this handler is called
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function registerUserSettingsChangeHandler(settingTypes, handler) {
    ensureInitialized(runtime);
    registerHandler('userSettingsChange', handler, true, [settingTypes]);
}
/**
 * @hidden
 * Opens a client-friendly preview of the specified file.
 *
 * @param file - The file to preview.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function openFilePreview(filePreviewParameters) {
    ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
    var params = [
        filePreviewParameters.entityId,
        filePreviewParameters.title,
        filePreviewParameters.description,
        filePreviewParameters.type,
        filePreviewParameters.objectUrl,
        filePreviewParameters.downloadUrl,
        filePreviewParameters.webPreviewUrl,
        filePreviewParameters.webEditUrl,
        filePreviewParameters.baseUrl,
        filePreviewParameters.editFile,
        filePreviewParameters.subEntityId,
        filePreviewParameters.viewerAction,
        filePreviewParameters.fileOpenPreference,
        filePreviewParameters.conversationId,
    ];
    sendMessageToParent('openFilePreview', params);
}

;// CONCATENATED MODULE: ./src/private/conversations.ts





/**
 * @hidden
 * Namespace to interact with the conversational subEntities inside the tab
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var conversations;
(function (conversations) {
    /**
     * @hidden
     * Hide from docs
     * --------------
     * Allows the user to start or continue a conversation with each subentity inside the tab
     *
     * @returns Promise resolved upon completion
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function openConversation(openConversationRequest) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            var sendPromise = sendAndHandleStatusAndReason('conversations.openConversation', {
                title: openConversationRequest.title,
                subEntityId: openConversationRequest.subEntityId,
                conversationId: openConversationRequest.conversationId,
                channelId: openConversationRequest.channelId,
                entityId: openConversationRequest.entityId,
            });
            if (openConversationRequest.onStartConversation) {
                registerHandler('startConversation', function (subEntityId, conversationId, channelId, entityId) {
                    return openConversationRequest.onStartConversation({
                        subEntityId: subEntityId,
                        conversationId: conversationId,
                        channelId: channelId,
                        entityId: entityId,
                    });
                });
            }
            if (openConversationRequest.onCloseConversation) {
                registerHandler('closeConversation', function (subEntityId, conversationId, channelId, entityId) {
                    return openConversationRequest.onCloseConversation({
                        subEntityId: subEntityId,
                        conversationId: conversationId,
                        channelId: channelId,
                        entityId: entityId,
                    });
                });
            }
            resolve(sendPromise);
        });
    }
    conversations.openConversation = openConversation;
    /**
     * @hidden
     *
     * Allows the user to close the conversation in the right pane
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function closeConversation() {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('conversations.closeConversation');
        removeHandler('startConversation');
        removeHandler('closeConversation');
    }
    conversations.closeConversation = closeConversation;
    /**
     * @hidden
     * Hide from docs
     * ------
     * Allows retrieval of information for all chat members.
     * NOTE: This value should be used only as a hint as to who the members are
     * and never as proof of membership in case your app is being hosted by a malicious party.
     *
     * @returns Promise resolved with information on all chat members
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getChatMembers() {
        return new Promise(function (resolve) {
            ensureInitialized(runtime);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            resolve(sendAndUnwrap('getChatMembers'));
        });
    }
    conversations.getChatMembers = getChatMembers;
    /**
     * Checks if the conversations capability is supported by the host
     * @returns boolean to represent whether conversations capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.conversations ? true : false;
    }
    conversations.isSupported = isSupported;
})(conversations || (conversations = {}));

;// CONCATENATED MODULE: ./src/internal/deepLinkConstants.ts
/**
 * App install dialog constants
 */
var teamsDeepLinkUrlPathForAppInstall = '/l/app/';
/**
 * Calendar constants
 */
var teamsDeepLinkUrlPathForCalendar = '/l/meeting/new';
var teamsDeepLinkAttendeesUrlParameterName = 'attendees';
var teamsDeepLinkStartTimeUrlParameterName = 'startTime';
var teamsDeepLinkEndTimeUrlParameterName = 'endTime';
var teamsDeepLinkSubjectUrlParameterName = 'subject';
var teamsDeepLinkContentUrlParameterName = 'content';
/**
 * Call constants
 */
var teamsDeepLinkUrlPathForCall = '/l/call/0/0';
var teamsDeepLinkSourceUrlParameterName = 'source';
var teamsDeepLinkWithVideoUrlParameterName = 'withVideo';
/**
 * Chat constants
 */
var teamsDeepLinkUrlPathForChat = '/l/chat/0/0';
var teamsDeepLinkUsersUrlParameterName = 'users';
var teamsDeepLinkTopicUrlParameterName = 'topicName';
var teamsDeepLinkMessageUrlParameterName = 'message';

;// CONCATENATED MODULE: ./src/internal/deepLinkUtilities.ts


function createTeamsDeepLinkForChat(users, topic, message) {
    if (users.length === 0) {
        throw new Error('Must have at least one user when creating a chat deep link');
    }
    var usersSearchParameter = "".concat(teamsDeepLinkUsersUrlParameterName, "=") + users.map(function (user) { return encodeURIComponent(user); }).join(',');
    var topicSearchParameter = topic === undefined ? '' : "&".concat(teamsDeepLinkTopicUrlParameterName, "=").concat(encodeURIComponent(topic));
    var messageSearchParameter = message === undefined ? '' : "&".concat(teamsDeepLinkMessageUrlParameterName, "=").concat(encodeURIComponent(message));
    return "".concat(teamsDeepLinkProtocol, "://").concat(teamsDeepLinkHost).concat(teamsDeepLinkUrlPathForChat, "?").concat(usersSearchParameter).concat(topicSearchParameter).concat(messageSearchParameter);
}
function createTeamsDeepLinkForCall(targets, withVideo, source) {
    if (targets.length === 0) {
        throw new Error('Must have at least one target when creating a call deep link');
    }
    var usersSearchParameter = "".concat(teamsDeepLinkUsersUrlParameterName, "=") + targets.map(function (user) { return encodeURIComponent(user); }).join(',');
    var withVideoSearchParameter = withVideo === undefined ? '' : "&".concat(teamsDeepLinkWithVideoUrlParameterName, "=").concat(encodeURIComponent(withVideo));
    var sourceSearchParameter = source === undefined ? '' : "&".concat(teamsDeepLinkSourceUrlParameterName, "=").concat(encodeURIComponent(source));
    return "".concat(teamsDeepLinkProtocol, "://").concat(teamsDeepLinkHost).concat(teamsDeepLinkUrlPathForCall, "?").concat(usersSearchParameter).concat(withVideoSearchParameter).concat(sourceSearchParameter);
}
function createTeamsDeepLinkForCalendar(attendees, startTime, endTime, subject, content) {
    var attendeeSearchParameter = attendees === undefined
        ? ''
        : "".concat(teamsDeepLinkAttendeesUrlParameterName, "=") +
            attendees.map(function (attendee) { return encodeURIComponent(attendee); }).join(',');
    var startTimeSearchParameter = startTime === undefined ? '' : "&".concat(teamsDeepLinkStartTimeUrlParameterName, "=").concat(encodeURIComponent(startTime));
    var endTimeSearchParameter = endTime === undefined ? '' : "&".concat(teamsDeepLinkEndTimeUrlParameterName, "=").concat(encodeURIComponent(endTime));
    var subjectSearchParameter = subject === undefined ? '' : "&".concat(teamsDeepLinkSubjectUrlParameterName, "=").concat(encodeURIComponent(subject));
    var contentSearchParameter = content === undefined ? '' : "&".concat(teamsDeepLinkContentUrlParameterName, "=").concat(encodeURIComponent(content));
    return "".concat(teamsDeepLinkProtocol, "://").concat(teamsDeepLinkHost).concat(teamsDeepLinkUrlPathForCalendar, "?").concat(attendeeSearchParameter).concat(startTimeSearchParameter).concat(endTimeSearchParameter).concat(subjectSearchParameter).concat(contentSearchParameter);
}
function createTeamsDeepLinkForAppInstallDialog(appId) {
    if (!appId) {
        throw new Error('App ID must be set when creating an app install dialog deep link');
    }
    return "".concat(teamsDeepLinkProtocol, "://").concat(teamsDeepLinkHost).concat(teamsDeepLinkUrlPathForAppInstall).concat(encodeURIComponent(appId));
}

;// CONCATENATED MODULE: ./src/public/appInstallDialog.ts






var appInstallDialog;
(function (appInstallDialog) {
    /**
     * Displays a dialog box that allows users to install a specific app within the host environment.
     *
     * @param openAPPInstallDialogParams - See {@link OpenAppInstallDialogParams | OpenAppInstallDialogParams} for more information.
     */
    function openAppInstallDialog(openAPPInstallDialogParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            if (runtime.isLegacyTeams) {
                resolve(sendAndHandleStatusAndReason('executeDeepLink', createTeamsDeepLinkForAppInstallDialog(openAPPInstallDialogParams.appId)));
            }
            else {
                sendMessageToParent('appInstallDialog.openAppInstallDialog', [openAPPInstallDialogParams]);
                resolve();
            }
        });
    }
    appInstallDialog.openAppInstallDialog = openAppInstallDialog;
    /**
     * Checks if the appInstallDialog capability is supported by the host
     * @returns boolean to represent whether the appInstallDialog capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.appInstallDialog ? true : false;
    }
    appInstallDialog.isSupported = isSupported;
})(appInstallDialog || (appInstallDialog = {}));

;// CONCATENATED MODULE: ./src/public/appNotification.ts




/**
 * Namespace to interact with the appNotification specific part of the SDK
 * @beta
 */
var appNotification;
(function (appNotification) {
    /**
     * @internal
     *
     * @hidden
     *
     * @beta
     *
     * This converts the notifcationActionUrl from a URL type to a string type for proper flow across the iframe
     * @param notificationDisplayParam - appNotification display parameter with the notificationActionUrl as a URL type
     * @returns a serialized object that can be sent to the host SDK
     */
    function serializeParam(notificationDisplayParam) {
        var _a;
        return {
            title: notificationDisplayParam.title,
            content: notificationDisplayParam.content,
            notificationIconAsSring: (_a = notificationDisplayParam.icon) === null || _a === void 0 ? void 0 : _a.href,
            displayDurationInSeconds: notificationDisplayParam.displayDurationInSeconds,
            notificationActionUrlAsString: notificationDisplayParam.notificationActionUrl.href,
        };
    }
    /**
     * Checks the valididty of the URL passed
     *
     * @param url
     * @returns True if a valid url was passed
     *
     * @beta
     */
    function isValidUrl(url) {
        var validProtocols = ['https:'];
        return validProtocols.includes(url.protocol);
    }
    /**
     * Validates that the input string is of property length
     *
     * @param inputString and maximumLength
     * @returns True if string length is within specified limit
     *
     * @beta
     */
    function isValidLength(inputString, maxLength) {
        return inputString.length <= maxLength;
    }
    /**
     * Validates that all the required appNotification display parameters are either supplied or satisfy the required checks
     * @param notificationDisplayParam
     * @throws Error if any of the required parameters aren't supplied
     * @throws Error if content or title length is beyond specific limit
     * @throws Error if invalid url is passed
     * @returns void
     *
     * @beta
     */
    function validateNotificationDisplayParams(notificationDisplayParam) {
        var maxTitleLength = 75;
        var maxContentLength = 1500;
        if (!notificationDisplayParam.title) {
            throw new Error('Must supply notification title');
        }
        if (!isValidLength(notificationDisplayParam.title, maxTitleLength)) {
            throw new Error("Invalid notification title length: Maximum title length ".concat(maxTitleLength, ", title length supplied ").concat(notificationDisplayParam.title.length));
        }
        if (!notificationDisplayParam.content) {
            throw new Error('Must supply notification content');
        }
        if (!isValidLength(notificationDisplayParam.content, maxContentLength)) {
            throw new Error("Invalid notification content length: Maximum content length ".concat(maxContentLength, ", content length supplied ").concat(notificationDisplayParam.content.length));
        }
        if (!notificationDisplayParam.notificationActionUrl) {
            throw new Error('Must supply notification url');
        }
        if (!isValidUrl(notificationDisplayParam.notificationActionUrl)) {
            throw new Error('Invalid notificationAction url');
        }
        if ((notificationDisplayParam === null || notificationDisplayParam === void 0 ? void 0 : notificationDisplayParam.icon) !== undefined && !isValidUrl(notificationDisplayParam === null || notificationDisplayParam === void 0 ? void 0 : notificationDisplayParam.icon)) {
            throw new Error('Invalid icon url');
        }
        if (!notificationDisplayParam.displayDurationInSeconds) {
            throw new Error('Must supply notification display duration in seconds');
        }
        if (notificationDisplayParam.displayDurationInSeconds < 0) {
            throw new Error('Notification display time must be greater than zero');
        }
    }
    /**
     * Displays appNotification after making a validiity check on all of the required parameters, by calling the validateNotificationDisplayParams helper function
     * An interface object containing all the required parameters to be displayed on the notification would be passed in here
     * The notificationDisplayParam would be serialized before passing across to the message handler to ensure all objects passed contain simple parameters that would properly pass across the Iframe
     * @param notificationdisplayparam - Interface object with all the parameters required to display an appNotificiation
     * @returns a promise resolution upon conclusion
     * @throws Error if appNotification capability is not supported
     * @throws Error if notficationDisplayParam was not validated successfully
     *
     * @beta
     */
    function displayInAppNotification(notificationDisplayParam) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.stage, FrameContexts.sidePanel, FrameContexts.meetingStage);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        validateNotificationDisplayParams(notificationDisplayParam);
        return sendAndHandleSdkError('appNotification.displayNotification', serializeParam(notificationDisplayParam));
    }
    appNotification.displayInAppNotification = displayInAppNotification;
    /**
     * Checks if appNotification is supported by the host
     * @returns boolean to represent whether the appNotification capability is supported
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.appNotification ? true : false;
    }
    appNotification.isSupported = isSupported;
})(appNotification || (appNotification = {}));

;// CONCATENATED MODULE: ./src/public/media.ts
/* eslint-disable @typescript-eslint/explicit-member-accessibility */
var __extends = (undefined && undefined.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();










/**
 * Interact with media, including capturing and viewing images.
 */
var media;
(function (media) {
    /**
     * Enum for file formats supported
     */
    var FileFormat;
    (function (FileFormat) {
        /** Base64 encoding */
        FileFormat["Base64"] = "base64";
        /** File id */
        FileFormat["ID"] = "id";
    })(FileFormat = media.FileFormat || (media.FileFormat = {}));
    /**
     * File object that can be used to represent image or video or audio
     */
    var File = /** @class */ (function () {
        function File() {
        }
        return File;
    }());
    media.File = File;
    /**
     * Launch camera, capture image or choose image from gallery and return the images as a File[] object to the callback.
     *
     * @params callback - Callback will be called with an @see SdkError if there are any.
     * If error is null or undefined, the callback will be called with a collection of @see File objects
     * @remarks
     * Note: Currently we support getting one File through this API, i.e. the file arrays size will be one.
     * Note: For desktop, this API is not supported. Callback will be resolved with ErrorCode.NotSupported.
     *
     */
    function captureImage(callback) {
        if (!callback) {
            throw new Error('[captureImage] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!GlobalVars.isFramelessWindow) {
            var notSupportedError = { errorCode: ErrorCode.NOT_SUPPORTED_ON_PLATFORM };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(notSupportedError, undefined);
            return;
        }
        if (!isCurrentSDKVersionAtLeast(captureImageMobileSupportVersion)) {
            var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(oldPlatformError, undefined);
            return;
        }
        sendMessageToParent('captureImage', callback);
    }
    media.captureImage = captureImage;
    /**
     * Checks whether or not media has user permission
     *
     * @returns Promise that will resolve with true if the user had granted the app permission to media information, or with false otherwise,
     * In case of an error, promise will reject with the error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
     *
     * @beta
     */
    function hasPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isPermissionSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.Media;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.has', permissions));
        });
    }
    media.hasPermission = hasPermission;
    /**
     * Requests user permission for media
     *
     * @returns Promise that will resolve with true if the user consented permission for media, or with false otherwise,
     * In case of an error, promise will reject with the error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
     *
     * @beta
     */
    function requestPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isPermissionSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.Media;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.request', permissions));
        });
    }
    media.requestPermission = requestPermission;
    /**
     * Checks if permission capability is supported by the host
     * @returns boolean to represent whether permission is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isPermissionSupported() {
        return ensureInitialized(runtime) && runtime.supports.permissions ? true : false;
    }
    /**
     * Media object returned by the select Media API
     */
    var Media = /** @class */ (function (_super) {
        __extends(Media, _super);
        function Media(that) {
            if (that === void 0) { that = null; }
            var _this = _super.call(this) || this;
            if (that) {
                _this.content = that.content;
                _this.format = that.format;
                _this.mimeType = that.mimeType;
                _this.name = that.name;
                _this.preview = that.preview;
                _this.size = that.size;
            }
            return _this;
        }
        /**
         * Gets the media in chunks irrespective of size, these chunks are assembled and sent back to the webapp as file/blob
         * @param callback - callback is called with the @see SdkError if there is an error
         * If error is null or undefined, the callback will be called with @see Blob.
         */
        Media.prototype.getMedia = function (callback) {
            if (!callback) {
                throw new Error('[get Media] Callback cannot be null');
            }
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isCurrentSDKVersionAtLeast(mediaAPISupportVersion)) {
                var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
                /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                callback(oldPlatformError, null);
                return;
            }
            if (!validateGetMediaInputs(this.mimeType, this.format, this.content)) {
                var invalidInput = { errorCode: ErrorCode.INVALID_ARGUMENTS };
                /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                callback(invalidInput, null);
                return;
            }
            // Call the new get media implementation via callbacks if the client version is greater than or equal to '2.0.0'
            if (isCurrentSDKVersionAtLeast(getMediaCallbackSupportVersion)) {
                this.getMediaViaCallback(callback);
            }
            else {
                this.getMediaViaHandler(callback);
            }
        };
        /** Function to retrieve media content, such as images or videos, via callback. */
        Media.prototype.getMediaViaCallback = function (callback) {
            var helper = {
                mediaMimeType: this.mimeType,
                assembleAttachment: [],
            };
            var localUriId = [this.content];
            function handleGetMediaCallbackRequest(mediaResult) {
                if (callback) {
                    if (mediaResult && mediaResult.error) {
                        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                        callback(mediaResult.error, null);
                    }
                    else {
                        if (mediaResult && mediaResult.mediaChunk) {
                            // If the chunksequence number is less than equal to 0 implies EOF
                            // create file/blob when all chunks have arrived and we get 0/-1 as chunksequence number
                            if (mediaResult.mediaChunk.chunkSequence <= 0) {
                                var file = createFile(helper.assembleAttachment, helper.mediaMimeType);
                                callback(mediaResult.error, file);
                            }
                            else {
                                // Keep pushing chunks into assemble attachment
                                var assemble = decodeAttachment(mediaResult.mediaChunk, helper.mediaMimeType);
                                helper.assembleAttachment.push(assemble);
                            }
                        }
                        else {
                            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                            callback({ errorCode: ErrorCode.INTERNAL_ERROR, message: 'data received is null' }, null);
                        }
                    }
                }
            }
            sendMessageToParent('getMedia', localUriId, handleGetMediaCallbackRequest);
        };
        /** Function to retrieve media content, such as images or videos, via handler. */
        Media.prototype.getMediaViaHandler = function (callback) {
            var actionName = generateGUID();
            var helper = {
                mediaMimeType: this.mimeType,
                assembleAttachment: [],
            };
            var params = [actionName, this.content];
            this.content && callback && sendMessageToParent('getMedia', params);
            function handleGetMediaRequest(response) {
                if (callback) {
                    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                    var mediaResult = JSON.parse(response);
                    if (mediaResult.error) {
                        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                        callback(mediaResult.error, null);
                        removeHandler('getMedia' + actionName);
                    }
                    else {
                        if (mediaResult.mediaChunk) {
                            // If the chunksequence number is less than equal to 0 implies EOF
                            // create file/blob when all chunks have arrived and we get 0/-1 as chunksequence number
                            if (mediaResult.mediaChunk.chunkSequence <= 0) {
                                var file = createFile(helper.assembleAttachment, helper.mediaMimeType);
                                callback(mediaResult.error, file);
                                removeHandler('getMedia' + actionName);
                            }
                            else {
                                // Keep pushing chunks into assemble attachment
                                var assemble = decodeAttachment(mediaResult.mediaChunk, helper.mediaMimeType);
                                helper.assembleAttachment.push(assemble);
                            }
                        }
                        else {
                            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                            callback({ errorCode: ErrorCode.INTERNAL_ERROR, message: 'data received is null' }, null);
                            removeHandler('getMedia' + actionName);
                        }
                    }
                }
            }
            registerHandler('getMedia' + actionName, handleGetMediaRequest);
        };
        return Media;
    }(File));
    media.Media = Media;
    /**
     * @hidden
     * Hide from docs
     * --------
     * Base class which holds the callback and notifies events to the host client
     */
    var MediaController = /** @class */ (function () {
        function MediaController(controllerCallback) {
            this.controllerCallback = controllerCallback;
        }
        /**
         * @hidden
         * Hide from docs
         * --------
         * Function to notify the host client to programatically control the experience
         * @param mediaEvent indicates what the event that needs to be signaled to the host client
         * Optional; @param callback is used to send app if host client has successfully handled the notification event or not
         */
        MediaController.prototype.notifyEventToHost = function (mediaEvent, callback) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            try {
                throwExceptionIfMobileApiIsNotSupported(nonFullScreenVideoModeAPISupportVersion);
            }
            catch (err) {
                if (callback) {
                    callback(err);
                }
                return;
            }
            var params = { mediaType: this.getMediaType(), mediaControllerEvent: mediaEvent };
            sendMessageToParent('media.controller', [params], function (err) {
                if (callback) {
                    callback(err);
                }
            });
        };
        /**
         * Function to programatically stop the ongoing media event
         * Optional; @param callback is used to send app if host client has successfully stopped the event or not
         */
        MediaController.prototype.stop = function (callback) {
            this.notifyEventToHost(MediaControllerEvent.StopRecording, callback);
        };
        return MediaController;
    }());
    /**
     * VideoController class is used to communicate between the app and the host client during the video capture flow
     */
    var VideoController = /** @class */ (function (_super) {
        __extends(VideoController, _super);
        function VideoController() {
            return _super !== null && _super.apply(this, arguments) || this;
        }
        /** Gets media type video. */
        VideoController.prototype.getMediaType = function () {
            return MediaType.Video;
        };
        /** Notify or send an event related to the playback and control of video content to a registered application. */
        VideoController.prototype.notifyEventToApp = function (mediaEvent) {
            if (!this.controllerCallback) {
                // Early return as app has not registered with the callback
                return;
            }
            switch (mediaEvent) {
                case MediaControllerEvent.StartRecording:
                    if (this.controllerCallback.onRecordingStarted) {
                        this.controllerCallback.onRecordingStarted();
                        break;
                    }
            }
        };
        return VideoController;
    }(MediaController));
    media.VideoController = VideoController;
    /**
     * @beta
     * Events which are used to communicate between the app and the host client during the media recording flow
     */
    var MediaControllerEvent;
    (function (MediaControllerEvent) {
        /** Start recording. */
        MediaControllerEvent[MediaControllerEvent["StartRecording"] = 1] = "StartRecording";
        /** Stop recording. */
        MediaControllerEvent[MediaControllerEvent["StopRecording"] = 2] = "StopRecording";
    })(MediaControllerEvent = media.MediaControllerEvent || (media.MediaControllerEvent = {}));
    /**
     * The modes in which camera can be launched in select Media API
     */
    var CameraStartMode;
    (function (CameraStartMode) {
        /** Photo mode. */
        CameraStartMode[CameraStartMode["Photo"] = 1] = "Photo";
        /** Document mode. */
        CameraStartMode[CameraStartMode["Document"] = 2] = "Document";
        /** Whiteboard mode. */
        CameraStartMode[CameraStartMode["Whiteboard"] = 3] = "Whiteboard";
        /** Business card mode. */
        CameraStartMode[CameraStartMode["BusinessCard"] = 4] = "BusinessCard";
    })(CameraStartMode = media.CameraStartMode || (media.CameraStartMode = {}));
    /**
     * Specifies the image source
     */
    var Source;
    (function (Source) {
        /** Image source is camera. */
        Source[Source["Camera"] = 1] = "Camera";
        /** Image source is gallery. */
        Source[Source["Gallery"] = 2] = "Gallery";
    })(Source = media.Source || (media.Source = {}));
    /**
     * Specifies the type of Media
     */
    var MediaType;
    (function (MediaType) {
        /** Media type photo or image */
        MediaType[MediaType["Image"] = 1] = "Image";
        /** Media type video. */
        MediaType[MediaType["Video"] = 2] = "Video";
        /** Media type video and image. */
        MediaType[MediaType["VideoAndImage"] = 3] = "VideoAndImage";
        /** Media type audio. */
        MediaType[MediaType["Audio"] = 4] = "Audio";
    })(MediaType = media.MediaType || (media.MediaType = {}));
    /**
     * ID contains a mapping for content uri on platform's side, URL is generic
     */
    var ImageUriType;
    (function (ImageUriType) {
        /** Image Id. */
        ImageUriType[ImageUriType["ID"] = 1] = "ID";
        /** Image URL. */
        ImageUriType[ImageUriType["URL"] = 2] = "URL";
    })(ImageUriType = media.ImageUriType || (media.ImageUriType = {}));
    /**
     * Specifies the image output formats.
     */
    var ImageOutputFormats;
    (function (ImageOutputFormats) {
        /** Outputs image.  */
        ImageOutputFormats[ImageOutputFormats["IMAGE"] = 1] = "IMAGE";
        /** Outputs pdf. */
        ImageOutputFormats[ImageOutputFormats["PDF"] = 2] = "PDF";
    })(ImageOutputFormats = media.ImageOutputFormats || (media.ImageOutputFormats = {}));
    /**
     * Select an attachment using camera/gallery
     *
     * @param mediaInputs - The input params to customize the media to be selected
     * @param callback - The callback to invoke after fetching the media
     */
    function selectMedia(mediaInputs, callback) {
        if (!callback) {
            throw new Error('[select Media] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isCurrentSDKVersionAtLeast(mediaAPISupportVersion)) {
            var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(oldPlatformError, null);
            return;
        }
        try {
            throwExceptionIfMediaCallIsNotSupportedOnMobile(mediaInputs);
        }
        catch (err) {
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(err, null);
            return;
        }
        if (!validateSelectMediaInputs(mediaInputs)) {
            var invalidInput = { errorCode: ErrorCode.INVALID_ARGUMENTS };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(invalidInput, null);
            return;
        }
        var params = [mediaInputs];
        // What comes back from native as attachments would just be objects and will be missing getMedia method on them
        sendMessageToParent('selectMedia', params, function (err, localAttachments, mediaEvent) {
            // MediaControllerEvent response is used to notify the app about events and is a partial response to selectMedia
            if (mediaEvent) {
                if (isVideoControllerRegistered(mediaInputs)) {
                    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                    mediaInputs.videoProps.videoController.notifyEventToApp(mediaEvent);
                }
                return;
            }
            // Media Attachments are final response to selectMedia
            if (!localAttachments) {
                /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                callback(err, null);
                return;
            }
            var mediaArray = [];
            for (var _i = 0, localAttachments_1 = localAttachments; _i < localAttachments_1.length; _i++) {
                var attachment = localAttachments_1[_i];
                mediaArray.push(new Media(attachment));
            }
            callback(err, mediaArray);
        });
    }
    media.selectMedia = selectMedia;
    /**
     * View images using native image viewer
     *
     * @param uriList - list of URIs for images to be viewed - can be content URI or server URL. Supports up to 10 Images in a single call
     * @param callback - returns back error if encountered, returns null in case of success
     */
    function viewImages(uriList, callback) {
        if (!callback) {
            throw new Error('[view images] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isCurrentSDKVersionAtLeast(mediaAPISupportVersion)) {
            var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
            callback(oldPlatformError);
            return;
        }
        if (!validateViewImagesInput(uriList)) {
            var invalidInput = { errorCode: ErrorCode.INVALID_ARGUMENTS };
            callback(invalidInput);
            return;
        }
        var params = [uriList];
        sendMessageToParent('viewImages', params, callback);
    }
    media.viewImages = viewImages;
    /**
     * @deprecated
     * As of 2.1.0, please use {@link barCode.scanBarCode barCode.scanBarCode(config?: BarCodeConfig): Promise\<string\>} instead.
  
     * Scan Barcode/QRcode using camera
     *
     * @remarks
     * Note: For desktop and web, this API is not supported. Callback will be resolved with ErrorCode.NotSupported.
     *
     * @param callback - callback to invoke after scanning the barcode
     * @param config - optional input configuration to customize the barcode scanning experience
     */
    function scanBarCode(callback, config) {
        if (!callback) {
            throw new Error('[media.scanBarCode] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (GlobalVars.hostClientType === HostClientType.desktop ||
            GlobalVars.hostClientType === HostClientType.web ||
            GlobalVars.hostClientType === HostClientType.rigel ||
            GlobalVars.hostClientType === HostClientType.teamsRoomsWindows ||
            GlobalVars.hostClientType === HostClientType.teamsRoomsAndroid ||
            GlobalVars.hostClientType === HostClientType.teamsPhones ||
            GlobalVars.hostClientType === HostClientType.teamsDisplays) {
            var notSupportedError = { errorCode: ErrorCode.NOT_SUPPORTED_ON_PLATFORM };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(notSupportedError, null);
            return;
        }
        if (!isCurrentSDKVersionAtLeast(scanBarCodeAPIMobileSupportVersion)) {
            var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(oldPlatformError, null);
            return;
        }
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        if (!validateScanBarCodeInput(config)) {
            var invalidInput = { errorCode: ErrorCode.INVALID_ARGUMENTS };
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            callback(invalidInput, null);
            return;
        }
        sendMessageToParent('media.scanBarCode', [config], callback);
    }
    media.scanBarCode = scanBarCode;
})(media || (media = {}));

;// CONCATENATED MODULE: ./src/internal/mediaUtil.ts



/**
 * @hidden
 * Helper function to create a blob from media chunks based on their sequence
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function createFile(assembleAttachment, mimeType) {
    if (assembleAttachment == null || mimeType == null || assembleAttachment.length <= 0) {
        return null;
    }
    var file;
    var sequence = 1;
    assembleAttachment.sort(function (a, b) { return (a.sequence > b.sequence ? 1 : -1); });
    assembleAttachment.forEach(function (item) {
        if (item.sequence == sequence) {
            if (file) {
                file = new Blob([file, item.file], { type: mimeType });
            }
            else {
                file = new Blob([item.file], { type: mimeType });
            }
            sequence++;
        }
    });
    return file;
}
/**
 * @hidden
 * Helper function to convert Media chunks into another object type which can be later assemebled
 * Converts base 64 encoded string to byte array and then into an array of blobs
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function decodeAttachment(attachment, mimeType) {
    if (attachment == null || mimeType == null) {
        return null;
    }
    var decoded = atob(attachment.chunk);
    var byteNumbers = new Array(decoded.length);
    for (var i = 0; i < decoded.length; i++) {
        byteNumbers[i] = decoded.charCodeAt(i);
    }
    var byteArray = new Uint8Array(byteNumbers);
    var blob = new Blob([byteArray], { type: mimeType });
    var assemble = {
        sequence: attachment.chunkSequence,
        file: blob,
    };
    return assemble;
}
/**
 * @hidden
 * Function throws an SdkError if the media call is not supported on current mobile version, else undefined.
 *
 * @throws an SdkError if the media call is not supported
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function throwExceptionIfMediaCallIsNotSupportedOnMobile(mediaInputs) {
    if (isMediaCallForVideoAndImageInputs(mediaInputs)) {
        throwExceptionIfMobileApiIsNotSupported(videoAndImageMediaAPISupportVersion);
    }
    else if (isMediaCallForNonFullScreenVideoMode(mediaInputs)) {
        throwExceptionIfMobileApiIsNotSupported(nonFullScreenVideoModeAPISupportVersion);
    }
    else if (isMediaCallForImageOutputFormats(mediaInputs)) {
        throwExceptionIfMobileApiIsNotSupported(imageOutputFormatsAPISupportVersion);
    }
}
/**
 * @hidden
 * Function returns true if the app has registered to listen to video controller events, else false.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function isVideoControllerRegistered(mediaInputs) {
    if (mediaInputs.mediaType == media.MediaType.Video &&
        mediaInputs.videoProps &&
        mediaInputs.videoProps.videoController) {
        return true;
    }
    return false;
}
/**
 * @hidden
 * Returns true if the mediaInput params are valid and false otherwise
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateSelectMediaInputs(mediaInputs) {
    if (mediaInputs == null || mediaInputs.maxMediaCount > 10) {
        return false;
    }
    return true;
}
/**
 * @hidden
 * Returns true if the mediaInput params are called for mediatype Image and contains Image outputs formats, false otherwise
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function isMediaCallForImageOutputFormats(mediaInputs) {
    var _a;
    if ((mediaInputs === null || mediaInputs === void 0 ? void 0 : mediaInputs.mediaType) == media.MediaType.Image && ((_a = mediaInputs === null || mediaInputs === void 0 ? void 0 : mediaInputs.imageProps) === null || _a === void 0 ? void 0 : _a.imageOutputFormats)) {
        return true;
    }
    return false;
}
/**
 * @hidden
 * Returns true if the mediaInput params are called for mediatype VideoAndImage and false otherwise
 *
 * @internal
 */
function isMediaCallForVideoAndImageInputs(mediaInputs) {
    if (mediaInputs && (mediaInputs.mediaType == media.MediaType.VideoAndImage || mediaInputs.videoAndImageProps)) {
        return true;
    }
    return false;
}
/**
 * @hidden
 * Returns true if the mediaInput params are called for non-full screen video mode and false otherwise
 *
 * @internal
 */
function isMediaCallForNonFullScreenVideoMode(mediaInputs) {
    if (mediaInputs &&
        mediaInputs.mediaType == media.MediaType.Video &&
        mediaInputs.videoProps &&
        !mediaInputs.videoProps.isFullScreenMode) {
        return true;
    }
    return false;
}
/**
 * @hidden
 * Returns true if the get Media params are valid and false otherwise
 *
 * @internal
 */
function validateGetMediaInputs(mimeType, format, content) {
    if (mimeType == null || format == null || format != media.FileFormat.ID || content == null) {
        return false;
    }
    return true;
}
/**
 * @hidden
 * Returns true if the view images param is valid and false otherwise
 *
 * @internal
 */
function validateViewImagesInput(uriList) {
    if (uriList == null || uriList.length <= 0 || uriList.length > 10) {
        return false;
    }
    return true;
}
/**
 * @hidden
 * Returns true if the scan barcode param is valid and false otherwise
 *
 * @internal
 */
function validateScanBarCodeInput(barCodeConfig) {
    if (barCodeConfig) {
        if (barCodeConfig.timeOutIntervalInSec === null ||
            barCodeConfig.timeOutIntervalInSec <= 0 ||
            barCodeConfig.timeOutIntervalInSec > 60) {
            return false;
        }
    }
    return true;
}
/**
 * @hidden
 * Returns true if the people picker params are valid and false otherwise
 *
 * @internal
 */
function validatePeoplePickerInput(peoplePickerInputs) {
    if (peoplePickerInputs) {
        if (peoplePickerInputs.title) {
            if (typeof peoplePickerInputs.title !== 'string') {
                return false;
            }
        }
        if (peoplePickerInputs.setSelected) {
            if (typeof peoplePickerInputs.setSelected !== 'object') {
                return false;
            }
        }
        if (peoplePickerInputs.openOrgWideSearchInChatOrChannel) {
            if (typeof peoplePickerInputs.openOrgWideSearchInChatOrChannel !== 'boolean') {
                return false;
            }
        }
        if (peoplePickerInputs.singleSelect) {
            if (typeof peoplePickerInputs.singleSelect !== 'boolean') {
                return false;
            }
        }
    }
    return true;
}

;// CONCATENATED MODULE: ./src/public/barCode.ts






/**
 * Namespace to interact with the barcode scanning-specific part of the SDK.
 *
 * @beta
 */
var barCode;
(function (barCode) {
    /**
     * Scan Barcode or QRcode using camera
     *
     * @param barCodeConfig - input configuration to customize the barcode scanning experience
     *
     * @returns a scanned code
     *
     * @beta
     */
    function scanBarCode(barCodeConfig) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (!validateScanBarCodeInput(barCodeConfig)) {
                throw { errorCode: ErrorCode.INVALID_ARGUMENTS };
            }
            resolve(sendAndHandleSdkError('media.scanBarCode', barCodeConfig));
        });
    }
    barCode.scanBarCode = scanBarCode;
    /**
     * Checks whether or not media has user permission
     *
     * @returns true if the user has granted the app permission to media information, false otherwise
     *
     * @beta
     */
    function hasPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.Media;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.has', permissions));
        });
    }
    barCode.hasPermission = hasPermission;
    /**
     * Requests user permission for media
     *
     * @returns true if the user has granted the app permission to the media, false otherwise
     *
     * @beta
     */
    function requestPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.Media;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.request', permissions));
        });
    }
    barCode.requestPermission = requestPermission;
    /**
     * Checks if barCode capability is supported by the host
     * @returns boolean to represent whether the barCode capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.barCode && runtime.supports.permissions ? true : false;
    }
    barCode.isSupported = isSupported;
})(barCode || (barCode = {}));

;// CONCATENATED MODULE: ./src/public/chat.ts





/**
 * Contains functionality to start chat with others
 *
 * @beta
 */
var chat;
(function (chat) {
    /**
     * Allows the user to open a chat with a single user and allows
     * for the user to specify the message they wish to send.
     *
     * @param openChatRequest: {@link OpenSingleChatRequest}- a request object that contains a user's email as well as an optional message parameter.
     *
     * @returns Promise resolved upon completion
     *
     * @beta
     */
    function openChat(openChatRequest) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (runtime.isLegacyTeams) {
                resolve(sendAndHandleStatusAndReason('executeDeepLink', createTeamsDeepLinkForChat([openChatRequest.user], undefined /*topic*/, openChatRequest.message)));
            }
            else {
                var sendPromise = sendAndHandleStatusAndReason('chat.openChat', {
                    members: openChatRequest.user,
                    message: openChatRequest.message,
                });
                resolve(sendPromise);
            }
        });
    }
    chat.openChat = openChat;
    /**
     * Allows the user to create a chat with multiple users (2+) and allows
     * for the user to specify a message and name the topic of the conversation. If
     * only 1 user is provided into users array default back to origin openChat.
     *
     * @param openChatRequest: {@link OpenGroupChatRequest} - a request object that contains a list of user emails as well as optional parameters for message and topic (display name for the group chat).
     *
     * @returns Promise resolved upon completion
     *
     * @beta
     */
    function openGroupChat(openChatRequest) {
        return new Promise(function (resolve) {
            if (openChatRequest.users.length < 1) {
                throw Error('OpenGroupChat Failed: No users specified');
            }
            if (openChatRequest.users.length === 1) {
                var chatRequest = {
                    user: openChatRequest.users[0],
                    message: openChatRequest.message,
                };
                openChat(chatRequest);
            }
            else {
                ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                if (runtime.isLegacyTeams) {
                    resolve(sendAndHandleStatusAndReason('executeDeepLink', createTeamsDeepLinkForChat(openChatRequest.users, openChatRequest.topic, openChatRequest.message)));
                }
                else {
                    var sendPromise = sendAndHandleStatusAndReason('chat.openChat', {
                        members: openChatRequest.users,
                        message: openChatRequest.message,
                        topic: openChatRequest.topic,
                    });
                    resolve(sendPromise);
                }
            }
        });
    }
    chat.openGroupChat = openGroupChat;
    /**
     * Checks if the chat capability is supported by the host
     * @returns boolean to represent whether the chat capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.chat ? true : false;
    }
    chat.isSupported = isSupported;
})(chat || (chat = {}));

;// CONCATENATED MODULE: ./src/public/geoLocation.ts





/**
 * Namespace to interact with the geoLocation module-specific part of the SDK. This is the newer version of location module.
 *
 * @beta
 */
var geoLocation;
(function (geoLocation) {
    /**
     * Fetches current user coordinates
     * @returns Promise that will resolve with {@link geoLocation.Location} object or reject with an error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
     *
     * @beta
     */
    function getCurrentLocation() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        return sendAndHandleSdkError('location.getLocation', { allowChooseLocation: false, showMap: false });
    }
    geoLocation.getCurrentLocation = getCurrentLocation;
    /**
     * Checks whether or not location has user permission
     *
     * @returns Promise that will resolve with true if the user had granted the app permission to location information, or with false otherwise,
     * In case of an error, promise will reject with the error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
     *
     * @beta
     */
    function hasPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.GeoLocation;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.has', permissions));
        });
    }
    geoLocation.hasPermission = hasPermission;
    /**
     * Requests user permission for location
     *
     * @returns true if the user consented permission for location, false otherwise
     * @returns Promise that will resolve with true if the user consented permission for location, or with false otherwise,
     * In case of an error, promise will reject with the error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
     *
     * @beta
     */
    function requestPermission() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        var permissions = DevicePermission.GeoLocation;
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('permissions.request', permissions));
        });
    }
    geoLocation.requestPermission = requestPermission;
    /**
     * Checks if geoLocation capability is supported by the host
     * @returns boolean to represent whether geoLocation is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.geoLocation && runtime.supports.permissions ? true : false;
    }
    geoLocation.isSupported = isSupported;
    /**
     * Namespace to interact with the location on map module-specific part of the SDK.
     *
     * @beta
     */
    var map;
    (function (map) {
        /**
         * Allows user to choose location on map
         *
         * @returns Promise that will resolve with {@link geoLocation.Location} object chosen by the user or reject with an error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
         *
         * @beta
         */
        function chooseLocation() {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            return sendAndHandleSdkError('location.getLocation', { allowChooseLocation: true, showMap: true });
        }
        map.chooseLocation = chooseLocation;
        /**
         * Shows the location on map corresponding to the given coordinates
         *
         * @param location - Location to be shown on the map
         * @returns Promise that resolves when the location dialog has been closed or reject with an error. Function can also throw a NOT_SUPPORTED_ON_PLATFORM error
         *
         * @beta
         */
        function showLocation(location) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (!location) {
                throw { errorCode: ErrorCode.INVALID_ARGUMENTS };
            }
            return sendAndHandleSdkError('location.showLocation', location);
        }
        map.showLocation = showLocation;
        /**
         * Checks if geoLocation.map capability is supported by the host
         * @returns boolean to represent whether geoLocation.map is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @beta
         */
        function isSupported() {
            return ensureInitialized(runtime) &&
                runtime.supports.geoLocation &&
                runtime.supports.geoLocation.map &&
                runtime.supports.permissions
                ? true
                : false;
        }
        map.isSupported = isSupported;
    })(map = geoLocation.map || (geoLocation.map = {}));
})(geoLocation || (geoLocation = {}));

;// CONCATENATED MODULE: ./src/public/adaptiveCards.ts

/**
 * @returns The {@linkcode AdaptiveCardVersion} representing the Adaptive Card schema
 * version supported by the host, or undefined if the host does not support Adaptive Cards
 */
function getAdaptiveCardSchemaVersion() {
    if (!runtime.hostVersionsInfo) {
        return undefined;
    }
    else {
        return runtime.hostVersionsInfo.adaptiveCardSchemaVersion;
    }
}

;// CONCATENATED MODULE: ./src/public/appWindow.ts
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/ban-types */






/**
 * An object that application can utilize to establish communication
 * with the child window it opened, which contains the corresponding task.
 */
var ChildAppWindow = /** @class */ (function () {
    function ChildAppWindow() {
    }
    /**
     * Send a message to the ChildAppWindow.
     *
     * @param message - The message to send
     * @param onComplete - The callback to know if the postMessage has been success/failed.
     */
    ChildAppWindow.prototype.postMessage = function (message, onComplete) {
        ensureInitialized(runtime);
        sendMessageToParent('messageForChild', [message], onComplete ? onComplete : getGenericOnCompleteHandler());
    };
    /**
     * Add a listener that will be called when an event is received from the ChildAppWindow.
     *
     * @param type - The event to listen to. Currently the only supported type is 'message'.
     * @param listener - The listener that will be called
     */
    ChildAppWindow.prototype.addEventListener = function (type, listener) {
        ensureInitialized(runtime);
        if (type === 'message') {
            registerHandler('messageForParent', listener);
        }
    };
    return ChildAppWindow;
}());

/**
 * An object that is utilized to facilitate communication with a parent window
 * that initiated the opening of current window. For instance, a dialog or task
 * module would utilize it to transmit messages to the application that launched it.
 */
var ParentAppWindow = /** @class */ (function () {
    function ParentAppWindow() {
    }
    Object.defineProperty(ParentAppWindow, "Instance", {
        /** Get the parent window instance. */
        get: function () {
            // Do you need arguments? Make it a regular method instead.
            return this._instance || (this._instance = new this());
        },
        enumerable: false,
        configurable: true
    });
    /**
     * Send a message to the ParentAppWindow.
     *
     * @param message - The message to send
     * @param onComplete - The callback to know if the postMessage has been success/failed.
     */
    ParentAppWindow.prototype.postMessage = function (message, onComplete) {
        ensureInitialized(runtime, FrameContexts.task);
        sendMessageToParent('messageForParent', [message], onComplete ? onComplete : getGenericOnCompleteHandler());
    };
    /**
     * Add a listener that will be called when an event is received from the ParentAppWindow.
     *
     * @param type - The event to listen to. Currently the only supported type is 'message'.
     * @param listener - The listener that will be called
     */
    ParentAppWindow.prototype.addEventListener = function (type, listener) {
        ensureInitialized(runtime, FrameContexts.task);
        if (type === 'message') {
            registerHandler('messageForChild', listener);
        }
    };
    return ParentAppWindow;
}());


;// CONCATENATED MODULE: ./src/public/secondaryBrowser.ts






/**
 * Namespace to power up the in-app browser experiences in the Host App.
 * For e.g., opening a URL in the Host App inside a browser
 *
 * @beta
 */
var secondaryBrowser;
(function (secondaryBrowser) {
    /**
     * Open a URL in the secondary browser aka in-app browser
     *
     * @param url Url to open in the browser
     * @returns Promise that successfully resolves if the URL  opens in the secondaryBrowser
     * or throws an error {@link SdkError} incase of failure before starting navigation
     *
     * @remarks Any error that happens after navigation begins is handled by the platform browser component and not returned from this function.
     * @beta
     */
    function open(url) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!url || !isValidHttpsURL(url)) {
            throw { errorCode: ErrorCode.INVALID_ARGUMENTS, message: 'Invalid Url: Only https URL is allowed' };
        }
        return sendAndHandleSdkError('secondaryBrowser.open', url.toString());
    }
    secondaryBrowser.open = open;
    /**
     * Checks if secondaryBrowser capability is supported by the host
     * @returns boolean to represent whether secondaryBrowser is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.secondaryBrowser ? true : false;
    }
    secondaryBrowser.isSupported = isSupported;
})(secondaryBrowser || (secondaryBrowser = {}));

;// CONCATENATED MODULE: ./src/public/location.ts






/**
 * @deprecated
 * As of 2.1.0, please use geoLocation namespace.
 *
 * Namespace to interact with the location module-specific part of the SDK.
 */
var location_location;
(function (location_1) {
    /**
     * @deprecated
     * As of 2.1.0, please use one of the following functions:
     * - {@link geoLocation.getCurrentLocation geoLocation.getCurrentLocation(): Promise\<Location\>} to get the current location.
     * - {@link geoLocation.map.chooseLocation geoLocation.map.chooseLocation(): Promise\<Location\>} to choose location on map.
     *
     * Fetches user location
     * @param props {@link LocationProps} - Specifying how the location request is handled
     * @param callback - Callback to invoke when current user location is fetched
     */
    function getLocation(props, callback) {
        if (!callback) {
            throw new Error('[location.getLocation] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isCurrentSDKVersionAtLeast(locationAPIsRequiredVersion)) {
            throw { errorCode: ErrorCode.OLD_PLATFORM };
        }
        if (!props) {
            throw { errorCode: ErrorCode.INVALID_ARGUMENTS };
        }
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('location.getLocation', [props], callback);
    }
    location_1.getLocation = getLocation;
    /**
     * @deprecated
     * As of 2.1.0, please use {@link geoLocation.map.showLocation geoLocation.map.showLocation(location: Location): Promise\<void\>} instead.
     *
     * Shows the location on map corresponding to the given coordinates
     *
     * @param location - Location to be shown on the map
     * @param callback - Callback to invoke when the location is opened on map
     */
    function showLocation(location, callback) {
        if (!callback) {
            throw new Error('[location.showLocation] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isCurrentSDKVersionAtLeast(locationAPIsRequiredVersion)) {
            throw { errorCode: ErrorCode.OLD_PLATFORM };
        }
        if (!location) {
            throw { errorCode: ErrorCode.INVALID_ARGUMENTS };
        }
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('location.showLocation', [location], callback);
    }
    location_1.showLocation = showLocation;
    /**
     * @deprecated
     * As of 2.1.0, please use geoLocation namespace, and use {@link geoLocation.isSupported geoLocation.isSupported: boolean} to check if geoLocation is supported.
     *
     * Checks if Location capability is supported by the host
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @returns boolean to represent whether Location is supported
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.location ? true : false;
    }
    location_1.isSupported = isSupported;
})(location_location || (location_location = {}));

;// CONCATENATED MODULE: ./src/public/meeting.ts
var __awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};





/**
 * Interact with meetings, including retrieving meeting details, getting mic status, and sharing app content.
 * This namespace is used to handle meeting related functionality like
 * get meeting details, get/update state of mic, sharing app content and more.
 */
var meeting;
(function (meeting) {
    /**
     * Reasons for the app's microphone state to change
     */
    var MicStateChangeReason;
    (function (MicStateChangeReason) {
        MicStateChangeReason[MicStateChangeReason["HostInitiated"] = 0] = "HostInitiated";
        MicStateChangeReason[MicStateChangeReason["AppInitiated"] = 1] = "AppInitiated";
        MicStateChangeReason[MicStateChangeReason["AppDeclinedToChange"] = 2] = "AppDeclinedToChange";
        MicStateChangeReason[MicStateChangeReason["AppFailedToChange"] = 3] = "AppFailedToChange";
    })(MicStateChangeReason || (MicStateChangeReason = {}));
    /**
     * Different types of meeting reactions that can be sent/received
     *
     * @hidden
     * Hide from docs.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    var MeetingReactionType;
    (function (MeetingReactionType) {
        MeetingReactionType["like"] = "like";
        MeetingReactionType["heart"] = "heart";
        MeetingReactionType["laugh"] = "laugh";
        MeetingReactionType["surprised"] = "surprised";
        MeetingReactionType["applause"] = "applause";
    })(MeetingReactionType = meeting.MeetingReactionType || (meeting.MeetingReactionType = {}));
    /** Represents the type of a meeting */
    var MeetingType;
    (function (MeetingType) {
        /** Used when the meeting type is not known. */
        MeetingType["Unknown"] = "Unknown";
        /** Used for ad hoc meetings that are created on the fly. */
        MeetingType["Adhoc"] = "Adhoc";
        /** Used for meetings that have been scheduled in advance. */
        MeetingType["Scheduled"] = "Scheduled";
        /** Used for meetings that occur on a recurring basis. */
        MeetingType["Recurring"] = "Recurring";
        /** Used for live events or webinars. */
        MeetingType["Broadcast"] = "Broadcast";
        /** Used for meetings that are created on the fly, but with a more polished experience than ad hoc meetings. */
        MeetingType["MeetNow"] = "MeetNow";
    })(MeetingType = meeting.MeetingType || (meeting.MeetingType = {}));
    /** Represents the type of a call. */
    var CallType;
    (function (CallType) {
        /** Represents a call between two people. */
        CallType["OneOnOneCall"] = "oneOnOneCall";
        /** Represents a call between more than two people. */
        CallType["GroupCall"] = "groupCall";
    })(CallType = meeting.CallType || (meeting.CallType = {}));
    /**
     * Allows an app to get the incoming audio speaker setting for the meeting user
     *
     * @param callback - Callback contains 2 parameters, error and result.
     *
     * error can either contain an error of type SdkError, incase of an error, or null when fetch is successful
     * result can either contain the true/false value, incase of a successful fetch or null when the fetching fails
     * result: True means incoming audio is muted and false means incoming audio is unmuted
     */
    function getIncomingClientAudioState(callback) {
        if (!callback) {
            throw new Error('[get incoming client audio state] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('getIncomingClientAudioState', callback);
    }
    meeting.getIncomingClientAudioState = getIncomingClientAudioState;
    /**
     * Allows an app to toggle the incoming audio speaker setting for the meeting user from mute to unmute or vice-versa
     *
     * @param callback - Callback contains 2 parameters, error and result.
     * error can either contain an error of type SdkError, incase of an error, or null when toggle is successful
     * result can either contain the true/false value, incase of a successful toggle or null when the toggling fails
     * result: True means incoming audio is muted and false means incoming audio is unmuted
     */
    function toggleIncomingClientAudio(callback) {
        if (!callback) {
            throw new Error('[toggle incoming client audio] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('toggleIncomingClientAudio', callback);
    }
    meeting.toggleIncomingClientAudio = toggleIncomingClientAudio;
    /**
     * @hidden
     * Allows an app to get the meeting details for the meeting
     *
     * @param callback - Callback contains 2 parameters, error and meetingDetailsResponse.
     * error can either contain an error of type SdkError, incase of an error, or null when get is successful
     * result can either contain a IMeetingDetailsResponse value, in case of a successful get or null when the get fails
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getMeetingDetails(callback) {
        if (!callback) {
            throw new Error('[get meeting details] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage, FrameContexts.settings, FrameContexts.content);
        sendMessageToParent('meeting.getMeetingDetails', callback);
    }
    meeting.getMeetingDetails = getMeetingDetails;
    /**
     * @hidden
     * Allows an app to get the authentication token for the anonymous or guest user in the meeting
     *
     * @param callback - Callback contains 2 parameters, error and authenticationTokenOfAnonymousUser.
     * error can either contain an error of type SdkError, incase of an error, or null when get is successful
     * authenticationTokenOfAnonymousUser can either contain a string value, incase of a successful get or null when the get fails
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getAuthenticationTokenForAnonymousUser(callback) {
        if (!callback) {
            throw new Error('[get Authentication Token For AnonymousUser] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage, FrameContexts.task);
        sendMessageToParent('meeting.getAuthenticationTokenForAnonymousUser', callback);
    }
    meeting.getAuthenticationTokenForAnonymousUser = getAuthenticationTokenForAnonymousUser;
    /**
     * Allows an app to get the state of the live stream in the current meeting
     *
     * @param callback - Callback contains 2 parameters: error and liveStreamState.
     * error can either contain an error of type SdkError, in case of an error, or null when get is successful
     * liveStreamState can either contain a LiveStreamState value, or null when operation fails
     */
    function getLiveStreamState(callback) {
        if (!callback) {
            throw new Error('[get live stream state] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        sendMessageToParent('meeting.getLiveStreamState', callback);
    }
    meeting.getLiveStreamState = getLiveStreamState;
    /**
     * Allows an app to request the live streaming be started at the given streaming url
     *
     * @remarks
     * Use getLiveStreamState or registerLiveStreamChangedHandler to get updates on the live stream state
     *
     * @param streamUrl - the url to the stream resource
     * @param streamKey - the key to the stream resource
     * @param callback - Callback contains error parameter which can be of type SdkError in case of an error, or null when operation is successful
     */
    function requestStartLiveStreaming(callback, streamUrl, streamKey) {
        if (!callback) {
            throw new Error('[request start live streaming] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        sendMessageToParent('meeting.requestStartLiveStreaming', [streamUrl, streamKey], callback);
    }
    meeting.requestStartLiveStreaming = requestStartLiveStreaming;
    /**
     * Allows an app to request the live streaming be stopped at the given streaming url
     *
     * @remarks
     * Use getLiveStreamState or registerLiveStreamChangedHandler to get updates on the live stream state
     *
     * @param callback - Callback contains error parameter which can be of type SdkError in case of an error, or null when operation is successful
     */
    function requestStopLiveStreaming(callback) {
        if (!callback) {
            throw new Error('[request stop live streaming] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        sendMessageToParent('meeting.requestStopLiveStreaming', callback);
    }
    meeting.requestStopLiveStreaming = requestStopLiveStreaming;
    /**
     * Registers a handler for changes to the live stream.
     *
     * @remarks
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the live stream state changes
     */
    function registerLiveStreamChangedHandler(handler) {
        if (!handler) {
            throw new Error('[register live stream changed handler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        registerHandler('meeting.liveStreamChanged', handler);
    }
    meeting.registerLiveStreamChangedHandler = registerLiveStreamChangedHandler;
    /**
     * Allows an app to share contents in the meeting
     *
     * @param callback - Callback contains 2 parameters, error and result.
     * error can either contain an error of type SdkError, incase of an error, or null when share is successful
     * result can either contain a true value, incase of a successful share or null when the share fails
     * @param appContentUrl - is the input URL which needs to be shared on to the stage
     */
    function shareAppContentToStage(callback, appContentUrl) {
        if (!callback) {
            throw new Error('[share app content to stage] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('meeting.shareAppContentToStage', [appContentUrl], callback);
    }
    meeting.shareAppContentToStage = shareAppContentToStage;
    /**
     * Provides information related app's in-meeting sharing capabilities
     *
     * @param callback - Callback contains 2 parameters, error and result.
     * error can either contain an error of type SdkError (error indication), or null (non-error indication)
     * appContentStageSharingCapabilities can either contain an IAppContentStageSharingCapabilities object
     * (indication of successful retrieval), or null (indication of failed retrieval)
     */
    function getAppContentStageSharingCapabilities(callback) {
        if (!callback) {
            throw new Error('[get app content stage sharing capabilities] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('meeting.getAppContentStageSharingCapabilities', callback);
    }
    meeting.getAppContentStageSharingCapabilities = getAppContentStageSharingCapabilities;
    /**
     * @hidden
     * Hide from docs.
     * Terminates current stage sharing session in meeting
     *
     * @param callback - Callback contains 2 parameters, error and result.
     * error can either contain an error of type SdkError (error indication), or null (non-error indication)
     * result can either contain a true boolean value (successful termination), or null (unsuccessful fetch)
     */
    function stopSharingAppContentToStage(callback) {
        if (!callback) {
            throw new Error('[stop sharing app content to stage] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('meeting.stopSharingAppContentToStage', callback);
    }
    meeting.stopSharingAppContentToStage = stopSharingAppContentToStage;
    /**
     * Provides information related to current stage sharing state for app
     *
     * @param callback - Callback contains 2 parameters, error and result.
     * error can either contain an error of type SdkError (error indication), or null (non-error indication)
     * appContentStageSharingState can either contain an IAppContentStageSharingState object
     * (indication of successful retrieval), or null (indication of failed retrieval)
     */
    function getAppContentStageSharingState(callback) {
        if (!callback) {
            throw new Error('[get app content stage sharing state] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('meeting.getAppContentStageSharingState', callback);
    }
    meeting.getAppContentStageSharingState = getAppContentStageSharingState;
    /**
     * Registers a handler for changes to paticipant speaking states. This API returns {@link ISpeakingState}, which will have isSpeakingDetected
     * and/or an error object. If any participant is speaking, isSpeakingDetected will be true. If no participants are speaking, isSpeakingDetected
     * will be false. Default value is false. Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler The handler to invoke when the speaking state of any participant changes (start/stop speaking).
     */
    function registerSpeakingStateChangeHandler(handler) {
        if (!handler) {
            throw new Error('[registerSpeakingStateChangeHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        registerHandler('meeting.speakingStateChanged', handler);
    }
    meeting.registerSpeakingStateChangeHandler = registerSpeakingStateChangeHandler;
    /**
     * Registers a handler for changes to the selfParticipant's (current user's) raiseHandState. If the selfParticipant raises their hand, isHandRaised
     * will be true. By default and if the selfParticipant hand is lowered, isHandRaised will be false. This API will return {@link RaiseHandStateChangedEventData}
     * that will have the raiseHandState or an error object. Only one handler can be registered at a time. A subsequent registration
     * replaces an existing registration.
     *
     * @param handler The handler to invoke when the selfParticipant's (current user's) raiseHandState changes.
     *
     * @hidden
     * Hide from docs.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    function registerRaiseHandStateChangedHandler(handler) {
        if (!handler) {
            throw new Error('[registerRaiseHandStateChangedHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        registerHandler('meeting.raiseHandStateChanged', handler);
    }
    meeting.registerRaiseHandStateChangedHandler = registerRaiseHandStateChangedHandler;
    /**
     * Registers a handler for receiving meeting reactions. When the selfParticipant (current user) successfully sends a meeting reaction and it is being rendered on the UI, the meetingReactionType will be populated. Only one handler can be registered
     * at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler The handler to invoke when the selfParticipant (current user) successfully sends a meeting reaction
     *
     * @hidden
     * Hide from docs.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    function registerMeetingReactionReceivedHandler(handler) {
        if (!handler) {
            throw new Error('[registerMeetingReactionReceivedHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        registerHandler('meeting.meetingReactionReceived', handler);
    }
    meeting.registerMeetingReactionReceivedHandler = registerMeetingReactionReceivedHandler;
    /**
     * Nested namespace for functions to control behavior of the app share button
     *
     * @beta
     */
    var appShareButton;
    (function (appShareButton) {
        /**
         * By default app share button will be hidden and this API will govern the visibility of it.
         *
         * This function can be used to hide/show app share button in meeting,
         * along with contentUrl (overrides contentUrl populated in app manifest)
         * @throws standard Invalid Url error
         * @param shareInformation has two elements, one isVisible boolean flag and another
         * optional string contentUrl, which will override contentUrl coming from Manifest
         * @beta
         */
        function setOptions(shareInformation) {
            ensureInitialized(runtime, FrameContexts.sidePanel);
            if (shareInformation.contentUrl) {
                new URL(shareInformation.contentUrl);
            }
            sendMessageToParent('meeting.appShareButton.setOptions', [shareInformation]);
        }
        appShareButton.setOptions = setOptions;
    })(appShareButton = meeting.appShareButton || (meeting.appShareButton = {}));
    /**
     * Have the app handle audio (mic & speaker) and turn off host audio.
     *
     * When {@link RequestAppAudioHandlingParams.isAppHandlingAudio} is true, the host will switch to audioless mode
     *   Registers for mic mute status change events, which are events that the app can receive from the host asking the app to
     *   mute or unmute the microphone.
     *
     * When {@link RequestAppAudioHandlingParams.isAppHandlingAudio} is false, the host will switch out of audioless mode
     *   Unregisters the mic mute status change events so the app will no longer receive these events
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     * @throws Error if {@link RequestAppAudioHandlingParams.micMuteStateChangedCallback} parameter is not defined
     *
     * @param requestAppAudioHandlingParams - {@link RequestAppAudioHandlingParams} object with values for the audio switchover
     * @param callback - Callback with one parameter, the result
     * can either be true (the host is now in audioless mode) or false (the host is not in audioless mode)
     *
     * @hidden
     * Hide from docs.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    function requestAppAudioHandling(requestAppAudioHandlingParams, callback) {
        if (!callback) {
            throw new Error('[requestAppAudioHandling] Callback response cannot be null');
        }
        if (!requestAppAudioHandlingParams.micMuteStateChangedCallback) {
            throw new Error('[requestAppAudioHandling] Callback Mic mute state handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        if (requestAppAudioHandlingParams.isAppHandlingAudio) {
            startAppAudioHandling(requestAppAudioHandlingParams, callback);
        }
        else {
            stopAppAudioHandling(requestAppAudioHandlingParams, callback);
        }
    }
    meeting.requestAppAudioHandling = requestAppAudioHandling;
    function startAppAudioHandling(requestAppAudioHandlingParams, callback) {
        var _this = this;
        var callbackInternalRequest = function (error, isHostAudioless) {
            if (error && isHostAudioless != null) {
                throw new Error('[requestAppAudioHandling] Callback response - both parameters cannot be set');
            }
            if (error) {
                throw new Error("[requestAppAudioHandling] Callback response - SDK error ".concat(error.errorCode, " ").concat(error.message));
            }
            if (typeof isHostAudioless !== 'boolean') {
                throw new Error('[requestAppAudioHandling] Callback response - isHostAudioless must be a boolean');
            }
            var micStateChangedCallback = function (micState) { return __awaiter(_this, void 0, void 0, function () {
                var newMicState, micStateDidUpdate, _a;
                return __generator(this, function (_b) {
                    switch (_b.label) {
                        case 0:
                            _b.trys.push([0, 2, , 3]);
                            return [4 /*yield*/, requestAppAudioHandlingParams.micMuteStateChangedCallback(micState)];
                        case 1:
                            newMicState = _b.sent();
                            micStateDidUpdate = newMicState.isMicMuted === micState.isMicMuted;
                            setMicStateWithReason(newMicState, micStateDidUpdate ? MicStateChangeReason.HostInitiated : MicStateChangeReason.AppDeclinedToChange);
                            return [3 /*break*/, 3];
                        case 2:
                            _a = _b.sent();
                            setMicStateWithReason(micState, MicStateChangeReason.AppFailedToChange);
                            return [3 /*break*/, 3];
                        case 3: return [2 /*return*/];
                    }
                });
            }); };
            registerHandler('meeting.micStateChanged', micStateChangedCallback);
            callback(isHostAudioless);
        };
        sendMessageToParent('meeting.requestAppAudioHandling', [requestAppAudioHandlingParams.isAppHandlingAudio], callbackInternalRequest);
    }
    function stopAppAudioHandling(requestAppAudioHandlingParams, callback) {
        var callbackInternalStop = function (error, isHostAudioless) {
            if (error && isHostAudioless != null) {
                throw new Error('[requestAppAudioHandling] Callback response - both parameters cannot be set');
            }
            if (error) {
                throw new Error("[requestAppAudioHandling] Callback response - SDK error ".concat(error.errorCode, " ").concat(error.message));
            }
            if (typeof isHostAudioless !== 'boolean') {
                throw new Error('[requestAppAudioHandling] Callback response - isHostAudioless must be a boolean');
            }
            if (doesHandlerExist('meeting.micStateChanged')) {
                removeHandler('meeting.micStateChanged');
            }
            callback(isHostAudioless);
        };
        sendMessageToParent('meeting.requestAppAudioHandling', [requestAppAudioHandlingParams.isAppHandlingAudio], callbackInternalStop);
    }
    /**
     * Notifies the host that the microphone state has changed in the app.
     * @param micState - The new state that the microphone is in
     *   isMicMuted - Boolean to indicate the current mute status of the mic.
     *
     * @hidden
     * Hide from docs.
     *
     * @internal
     * Limited to Microsoft-internal use
     *
     * @beta
     */
    function updateMicState(micState) {
        setMicStateWithReason(micState, MicStateChangeReason.AppInitiated);
    }
    meeting.updateMicState = updateMicState;
    function setMicStateWithReason(micState, reason) {
        ensureInitialized(runtime, FrameContexts.sidePanel, FrameContexts.meetingStage);
        sendMessageToParent('meeting.updateMicState', [micState, reason]);
    }
})(meeting || (meeting = {}));

;// CONCATENATED MODULE: ./src/public/monetization.ts





var monetization;
(function (monetization) {
    /**
     * @hidden
     * This function is the overloaded implementation of openPurchaseExperience.
     * Since the method signatures of the v1 callback and v2 promise differ in the type of the first parameter,
     * we need to do an extra check to know the typeof the @param1 to set the proper arguments of the utility function.
     * @param param1
     * @param param2
     * @returns Promise that will be resolved when the operation has completed or rejected with SdkError value
     */
    function openPurchaseExperience(param1, param2) {
        var callback;
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        var planInfo;
        if (typeof param1 === 'function') {
            callback = param1;
            planInfo = param2;
        }
        else {
            planInfo = param1;
        }
        var wrappedFunction = function () {
            return new Promise(function (resolve) {
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                resolve(sendAndHandleSdkError('monetization.openPurchaseExperience', planInfo));
            });
        };
        ensureInitialized(runtime, FrameContexts.content);
        return callCallbackWithErrorOrResultOrNullFromPromiseAndReturnPromise(wrappedFunction, callback);
    }
    monetization.openPurchaseExperience = openPurchaseExperience;
    /**
     * @hidden
     *
     * Checks if the monetization capability is supported by the host
     * @returns boolean to represent whether the monetization capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.monetization ? true : false;
    }
    monetization.isSupported = isSupported;
})(monetization || (monetization = {}));

;// CONCATENATED MODULE: ./src/public/calendar.ts





/**
 * Interact with the user's calendar, including opening calendar items and composing meetings.
 */
var calendar;
(function (calendar) {
    /**
     * Opens a calendar item.
     *
     * @param openCalendarItemParams - object containing unique ID of the calendar item to be opened.
     */
    function openCalendarItem(openCalendarItemParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            if (!openCalendarItemParams.itemId || !openCalendarItemParams.itemId.trim()) {
                throw new Error('Must supply an itemId to openCalendarItem');
            }
            resolve(sendAndHandleStatusAndReason('calendar.openCalendarItem', openCalendarItemParams));
        });
    }
    calendar.openCalendarItem = openCalendarItem;
    /**
     * Compose a new meeting in the user's calendar.
     *
     * @param composeMeetingParams - object containing various properties to set up the meeting details.
     */
    function composeMeeting(composeMeetingParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            if (runtime.isLegacyTeams) {
                resolve(sendAndHandleStatusAndReason('executeDeepLink', createTeamsDeepLinkForCalendar(composeMeetingParams.attendees, composeMeetingParams.startTime, composeMeetingParams.endTime, composeMeetingParams.subject, composeMeetingParams.content)));
            }
            else {
                resolve(sendAndHandleStatusAndReason('calendar.composeMeeting', composeMeetingParams));
            }
        });
    }
    calendar.composeMeeting = composeMeeting;
    /**
     * Checks if the calendar capability is supported by the host
     * @returns boolean to represent whether the calendar capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.calendar ? true : false;
    }
    calendar.isSupported = isSupported;
})(calendar || (calendar = {}));

;// CONCATENATED MODULE: ./src/public/mail.ts




/**
 * Used to interact with mail capability, including opening and composing mail.
 */
var mail;
(function (mail) {
    /**
     * Opens a mail message in the host.
     *
     * @param openMailItemParams - Object that specifies the ID of the mail message.
     */
    function openMailItem(openMailItemParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            if (!openMailItemParams.itemId || !openMailItemParams.itemId.trim()) {
                throw new Error('Must supply an itemId to openMailItem');
            }
            resolve(sendAndHandleStatusAndReason('mail.openMailItem', openMailItemParams));
        });
    }
    mail.openMailItem = openMailItem;
    /**
     * Compose a new email in the user's mailbox.
     *
     * @param composeMailParams - Object that specifies the type of mail item to compose and the details of the mail item.
     *
     */
    function composeMail(composeMailParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            resolve(sendAndHandleStatusAndReason('mail.composeMail', composeMailParams));
        });
    }
    mail.composeMail = composeMail;
    /**
     * Checks if the mail capability is supported by the host
     * @returns boolean to represent whether the mail capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.mail ? true : false;
    }
    mail.isSupported = isSupported;
    /** Defines compose mail types. */
    var ComposeMailType;
    (function (ComposeMailType) {
        /** Compose a new mail message. */
        ComposeMailType["New"] = "new";
        /** Compose a reply to the sender of an existing mail message. */
        ComposeMailType["Reply"] = "reply";
        /** Compose a reply to all recipients of an existing mail message. */
        ComposeMailType["ReplyAll"] = "replyAll";
        /** Compose a new mail message with the content of an existing mail message forwarded to a new recipient. */
        ComposeMailType["Forward"] = "forward";
    })(ComposeMailType = mail.ComposeMailType || (mail.ComposeMailType = {}));
})(mail || (mail = {}));

;// CONCATENATED MODULE: ./src/public/people.ts








var people;
(function (people_1) {
    /**
     * @hidden
     * This function is the overloaded implementation of selectPeople.
     * Since the method signatures of the v1 callback and v2 promise differ in the type of the first parameter,
     * we need to do an extra check to know the typeof the @param1 to set the proper arguments of the utility function.
     * @param param1
     * @param param2
     * @returns Promise of Array of PeoplePickerResult objects.
     */
    function selectPeople(param1, param2) {
        var _a;
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task, FrameContexts.settings);
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        var callback;
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        var peoplePickerInputs;
        if (typeof param1 === 'function') {
            _a = [param1, param2], callback = _a[0], peoplePickerInputs = _a[1];
        }
        else {
            peoplePickerInputs = param1;
        }
        return callCallbackWithErrorOrResultFromPromiseAndReturnPromise(selectPeopleHelper, callback, peoplePickerInputs);
    }
    people_1.selectPeople = selectPeople;
    function selectPeopleHelper(peoplePickerInputs) {
        return new Promise(function (resolve) {
            if (!isCurrentSDKVersionAtLeast(peoplePickerRequiredVersion)) {
                throw { errorCode: ErrorCode.OLD_PLATFORM };
            }
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            if (!validatePeoplePickerInput(peoplePickerInputs)) {
                throw { errorCode: ErrorCode.INVALID_ARGUMENTS };
            }
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
            resolve(sendAndHandleSdkError('people.selectPeople', peoplePickerInputs));
        });
    }
    /**
     * Checks if the people capability is supported by the host
     * @returns boolean to represent whether the people capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.people ? true : false;
    }
    people_1.isSupported = isSupported;
})(people || (people = {}));

;// CONCATENATED MODULE: ./src/internal/profileUtil.ts
/**
 * @hidden
 * Validates the request parameters
 * @param showProfileRequest The request parameters
 * @returns true if the parameters are valid, false otherwise
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateShowProfileRequest(showProfileRequest) {
    if (!showProfileRequest) {
        return [false, 'A request object is required'];
    }
    // Validate modality
    if (showProfileRequest.modality && typeof showProfileRequest.modality !== 'string') {
        return [false, 'modality must be a string'];
    }
    // Validate targetElementBoundingRect
    if (!showProfileRequest.targetElementBoundingRect ||
        typeof showProfileRequest.targetElementBoundingRect !== 'object') {
        return [false, 'targetElementBoundingRect must be a DOMRect'];
    }
    // Validate triggerType
    if (!showProfileRequest.triggerType || typeof showProfileRequest.triggerType !== 'string') {
        return [false, 'triggerType must be a valid string'];
    }
    return validatePersona(showProfileRequest.persona);
}
/**
 * @hidden
 * Validates the persona that is used to resolve the profile target
 * @param persona The persona object
 * @returns true if the persona is valid, false otherwise
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validatePersona(persona) {
    if (!persona) {
        return [false, 'persona object must be provided'];
    }
    if (persona.displayName && typeof persona.displayName !== 'string') {
        return [false, 'displayName must be a string'];
    }
    if (!persona.identifiers || typeof persona.identifiers !== 'object') {
        return [false, 'persona identifiers object must be provided'];
    }
    if (!persona.identifiers.AadObjectId && !persona.identifiers.Smtp && !persona.identifiers.Upn) {
        return [false, 'at least one valid identifier must be provided'];
    }
    if (persona.identifiers.AadObjectId && typeof persona.identifiers.AadObjectId !== 'string') {
        return [false, 'AadObjectId identifier must be a string'];
    }
    if (persona.identifiers.Smtp && typeof persona.identifiers.Smtp !== 'string') {
        return [false, 'Smtp identifier must be a string'];
    }
    if (persona.identifiers.Upn && typeof persona.identifiers.Upn !== 'string') {
        return [false, 'Upn identifier must be a string'];
    }
    return [true, undefined];
}

;// CONCATENATED MODULE: ./src/public/profile.ts






/**
 * Namespace for profile related APIs.
 *
 * @beta
 */
var profile;
(function (profile) {
    /**
     * Opens a profile card at a specified position to show profile information about a persona.
     * @param showProfileRequest The parameters to position the card and identify the target user.
     * @returns Promise that will be fulfilled when the operation has completed
     *
     * @beta
     */
    function showProfile(showProfileRequest) {
        ensureInitialized(runtime, FrameContexts.content);
        return new Promise(function (resolve) {
            var _a = validateShowProfileRequest(showProfileRequest), isValid = _a[0], message = _a[1];
            if (!isValid) {
                throw { errorCode: ErrorCode.INVALID_ARGUMENTS, message: message };
            }
            // Convert the app provided parameters to the form suitable for postMessage.
            var requestInternal = {
                modality: showProfileRequest.modality,
                persona: showProfileRequest.persona,
                triggerType: showProfileRequest.triggerType,
                targetRectangle: {
                    x: showProfileRequest.targetElementBoundingRect.x,
                    y: showProfileRequest.targetElementBoundingRect.y,
                    width: showProfileRequest.targetElementBoundingRect.width,
                    height: showProfileRequest.targetElementBoundingRect.height,
                },
            };
            resolve(sendAndHandleSdkError('profile.showProfile', requestInternal));
        });
    }
    profile.showProfile = showProfile;
    /**
     * Checks if the profile capability is supported by the host
     * @returns boolean to represent whether the profile capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.profile ? true : false;
    }
    profile.isSupported = isSupported;
})(profile || (profile = {}));

;// CONCATENATED MODULE: ./src/internal/videoFrameTick.ts

var VideoFrameTick = /** @class */ (function () {
    function VideoFrameTick() {
    }
    VideoFrameTick.setTimeout = function (callback, timeoutInMs) {
        var startedAtInMs = performance.now();
        var id = generateGUID();
        VideoFrameTick.setTimeoutCallbacks[id] = {
            callback: callback,
            timeoutInMs: timeoutInMs,
            startedAtInMs: startedAtInMs,
        };
        return id;
    };
    VideoFrameTick.clearTimeout = function (id) {
        delete VideoFrameTick.setTimeoutCallbacks[id];
    };
    VideoFrameTick.setInterval = function (callback, intervalInMs) {
        VideoFrameTick.setTimeout(function next() {
            callback();
            VideoFrameTick.setTimeout(next, intervalInMs);
        }, intervalInMs);
    };
    /**
     * Call this function whenever a frame comes in, it will check if any timeout is due and call the callback
     */
    VideoFrameTick.tick = function () {
        var now = performance.now();
        var timeoutIds = [];
        // find all the timeouts that are due,
        // not to invoke them in the loop to avoid modifying the collection while iterating
        for (var key in VideoFrameTick.setTimeoutCallbacks) {
            var callback = VideoFrameTick.setTimeoutCallbacks[key];
            var start = callback.startedAtInMs;
            if (now - start >= callback.timeoutInMs) {
                timeoutIds.push(key);
            }
        }
        // invoke the callbacks
        for (var _i = 0, timeoutIds_1 = timeoutIds; _i < timeoutIds_1.length; _i++) {
            var id = timeoutIds_1[_i];
            var callback = VideoFrameTick.setTimeoutCallbacks[id];
            callback.callback();
            delete VideoFrameTick.setTimeoutCallbacks[id];
        }
    };
    VideoFrameTick.setTimeoutCallbacks = {};
    return VideoFrameTick;
}());


;// CONCATENATED MODULE: ./src/internal/videoPerformanceStatistics.ts

var VideoPerformanceStatistics = /** @class */ (function () {
    function VideoPerformanceStatistics(distributionBinSize, 
    /**
     * Function to report the statistics result
     */
    reportStatisticsResult) {
        this.reportStatisticsResult = reportStatisticsResult;
        this.sampleCount = 0;
        this.distributionBins = new Uint32Array(distributionBinSize);
    }
    /**
     * Call this function before processing every frame
     */
    VideoPerformanceStatistics.prototype.processStarts = function (effectId, frameWidth, frameHeight, effectParam) {
        VideoFrameTick.tick();
        if (!this.suitableForThisSession(effectId, frameWidth, frameHeight, effectParam)) {
            this.reportAndResetSession(this.getStatistics(), effectId, effectParam, frameWidth, frameHeight);
        }
        this.start();
    };
    VideoPerformanceStatistics.prototype.processEnds = function () {
        // calculate duration of the process and record it
        var durationInMillisecond = performance.now() - this.frameProcessingStartedAt;
        var binIndex = Math.floor(Math.max(0, Math.min(this.distributionBins.length - 1, durationInMillisecond)));
        this.distributionBins[binIndex] += 1;
        this.sampleCount += 1;
    };
    VideoPerformanceStatistics.prototype.getStatistics = function () {
        if (!this.currentSession) {
            return null;
        }
        return {
            effectId: this.currentSession.effectId,
            effectParam: this.currentSession.effectParam,
            frameHeight: this.currentSession.frameHeight,
            frameWidth: this.currentSession.frameWidth,
            duration: performance.now() - this.currentSession.startedAtInMs,
            sampleCount: this.sampleCount,
            distributionBins: this.distributionBins.slice(),
        };
    };
    VideoPerformanceStatistics.prototype.start = function () {
        this.frameProcessingStartedAt = performance.now();
    };
    VideoPerformanceStatistics.prototype.suitableForThisSession = function (effectId, frameWidth, frameHeight, effectParam) {
        return (this.currentSession &&
            this.currentSession.effectId === effectId &&
            this.currentSession.effectParam === effectParam &&
            this.currentSession.frameWidth === frameWidth &&
            this.currentSession.frameHeight === frameHeight);
    };
    VideoPerformanceStatistics.prototype.reportAndResetSession = function (result, effectId, effectParam, frameWidth, frameHeight) {
        var _this = this;
        result && this.reportStatisticsResult(result);
        this.resetCurrentSession(this.getNextTimeout(effectId, this.currentSession), effectId, effectParam, frameWidth, frameHeight);
        if (this.timeoutId) {
            VideoFrameTick.clearTimeout(this.timeoutId);
        }
        this.timeoutId = VideoFrameTick.setTimeout((function () { return _this.reportAndResetSession(_this.getStatistics(), effectId, effectParam, frameWidth, frameHeight); }).bind(this), this.currentSession.timeoutInMs);
    };
    VideoPerformanceStatistics.prototype.resetCurrentSession = function (timeoutInMs, effectId, effectParam, frameWidth, frameHeight) {
        this.currentSession = {
            startedAtInMs: performance.now(),
            timeoutInMs: timeoutInMs,
            effectId: effectId,
            effectParam: effectParam,
            frameWidth: frameWidth,
            frameHeight: frameHeight,
        };
        this.sampleCount = 0;
        this.distributionBins.fill(0);
    };
    // send the statistics result every n second, where n starts from 1, 2, 4...and finally stays at every 30 seconds.
    VideoPerformanceStatistics.prototype.getNextTimeout = function (effectId, currentSession) {
        // only reset timeout when new session or effect changed
        if (!currentSession || currentSession.effectId !== effectId) {
            return VideoPerformanceStatistics.initialSessionTimeoutInMs;
        }
        return Math.min(VideoPerformanceStatistics.maxSessionTimeoutInMs, currentSession.timeoutInMs * 2);
    };
    VideoPerformanceStatistics.initialSessionTimeoutInMs = 1000;
    VideoPerformanceStatistics.maxSessionTimeoutInMs = 1000 * 30;
    return VideoPerformanceStatistics;
}());


;// CONCATENATED MODULE: ./src/internal/videoPerformanceMonitor.ts


/**
 * This class is used to monitor the performance of video processing, and report performance events.
 */
var VideoPerformanceMonitor = /** @class */ (function () {
    function VideoPerformanceMonitor(reportPerformanceEvent) {
        var _this = this;
        this.reportPerformanceEvent = reportPerformanceEvent;
        this.isFirstFrameProcessed = false;
        this.frameProcessTimeLimit = 100;
        this.frameProcessingStartedAt = 0;
        this.frameProcessingTimeCost = 0;
        this.processedFrameCount = 0;
        this.performanceStatistics = new VideoPerformanceStatistics(VideoPerformanceMonitor.distributionBinSize, function (result) {
            return _this.reportPerformanceEvent('video.performance.performanceDataGenerated', [result]);
        });
    }
    /**
     * Start to check frame processing time intervally
     * and report performance event if the average frame processing time is too long.
     */
    VideoPerformanceMonitor.prototype.startMonitorSlowFrameProcessing = function () {
        var _this = this;
        VideoFrameTick.setInterval(function () {
            if (_this.processedFrameCount === 0) {
                return;
            }
            var averageFrameProcessingTime = _this.frameProcessingTimeCost / _this.processedFrameCount;
            if (averageFrameProcessingTime > _this.frameProcessTimeLimit) {
                _this.reportPerformanceEvent('video.performance.frameProcessingSlow', [averageFrameProcessingTime]);
            }
            _this.frameProcessingTimeCost = 0;
            _this.processedFrameCount = 0;
        }, VideoPerformanceMonitor.calculateFPSInterval);
    };
    /**
     * Define the time limit of frame processing.
     * When the average frame processing time is longer than the time limit, a "video.performance.frameProcessingSlow" event will be reported.
     * @param timeLimit
     */
    VideoPerformanceMonitor.prototype.setFrameProcessTimeLimit = function (timeLimit) {
        this.frameProcessTimeLimit = timeLimit;
    };
    /**
     * Call this function when the app starts to switch to the new video effect
     */
    VideoPerformanceMonitor.prototype.reportApplyingVideoEffect = function (effectId, effectParam) {
        var _a, _b;
        if (((_a = this.applyingEffect) === null || _a === void 0 ? void 0 : _a.effectId) === effectId && ((_b = this.applyingEffect) === null || _b === void 0 ? void 0 : _b.effectParam) === effectParam) {
            return;
        }
        this.applyingEffect = {
            effectId: effectId,
            effectParam: effectParam,
        };
        this.appliedEffect = undefined;
    };
    /**
     * Call this function when the new video effect is ready
     */
    VideoPerformanceMonitor.prototype.reportVideoEffectChanged = function (effectId, effectParam) {
        if (this.applyingEffect === undefined ||
            (this.applyingEffect.effectId !== effectId && this.applyingEffect.effectParam !== effectParam)) {
            // don't handle obsoleted event
            return;
        }
        this.appliedEffect = {
            effectId: effectId,
            effectParam: effectParam,
        };
        this.applyingEffect = undefined;
        this.isFirstFrameProcessed = false;
    };
    /**
     * Call this function when the app starts to process a video frame
     */
    VideoPerformanceMonitor.prototype.reportStartFrameProcessing = function (frameWidth, frameHeight) {
        VideoFrameTick.tick();
        if (!this.appliedEffect) {
            return;
        }
        this.frameProcessingStartedAt = performance.now();
        this.performanceStatistics.processStarts(this.appliedEffect.effectId, frameWidth, frameHeight, this.appliedEffect.effectParam);
    };
    /**
     * Call this function when the app finishes successfully processing a video frame
     */
    VideoPerformanceMonitor.prototype.reportFrameProcessed = function () {
        var _a;
        if (!this.appliedEffect) {
            return;
        }
        this.processedFrameCount++;
        this.frameProcessingTimeCost += performance.now() - this.frameProcessingStartedAt;
        this.performanceStatistics.processEnds();
        if (!this.isFirstFrameProcessed) {
            this.isFirstFrameProcessed = true;
            this.reportPerformanceEvent('video.performance.firstFrameProcessed', [
                Date.now(),
                this.appliedEffect.effectId,
                (_a = this.appliedEffect) === null || _a === void 0 ? void 0 : _a.effectParam,
            ]);
        }
    };
    /**
     * Call this function when the app starts to get the texture stream
     */
    VideoPerformanceMonitor.prototype.reportGettingTextureStream = function (streamId) {
        this.gettingTextureStreamStartedAt = performance.now();
        this.currentStreamId = streamId;
    };
    /**
     * Call this function when the app finishes successfully getting the texture stream
     */
    VideoPerformanceMonitor.prototype.reportTextureStreamAcquired = function () {
        if (this.gettingTextureStreamStartedAt !== undefined) {
            var timeTaken = performance.now() - this.gettingTextureStreamStartedAt;
            this.reportPerformanceEvent('video.performance.textureStreamAcquired', [this.currentStreamId, timeTaken]);
        }
    };
    VideoPerformanceMonitor.distributionBinSize = 1000;
    VideoPerformanceMonitor.calculateFPSInterval = 1000;
    return VideoPerformanceMonitor;
}());


;// CONCATENATED MODULE: ./src/internal/videoUtils.ts
var videoUtils_awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var videoUtils_generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};





/**
 * @hidden
 * Create a MediaStreamTrack from the media stream with the given streamId and processed by videoFrameHandler.
 */
function processMediaStream(streamId, videoFrameHandler, notifyError, videoPerformanceMonitor) {
    return videoUtils_awaiter(this, void 0, void 0, function () {
        var _a;
        return videoUtils_generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _a = createProcessedStreamGenerator;
                    return [4 /*yield*/, getInputVideoTrack(streamId, notifyError, videoPerformanceMonitor)];
                case 1: return [2 /*return*/, _a.apply(void 0, [_b.sent(), new DefaultTransformer(notifyError, videoFrameHandler)])];
            }
        });
    });
}
/**
 * @hidden
 * Create a MediaStreamTrack from the media stream with the given streamId and processed by videoFrameHandler.
 * The videoFrameHandler will receive metadata of the video frame.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function processMediaStreamWithMetadata(streamId, videoFrameHandler, notifyError, videoPerformanceMonitor) {
    return videoUtils_awaiter(this, void 0, void 0, function () {
        var _a;
        return videoUtils_generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _a = createProcessedStreamGenerator;
                    return [4 /*yield*/, getInputVideoTrack(streamId, notifyError, videoPerformanceMonitor)];
                case 1: return [2 /*return*/, _a.apply(void 0, [_b.sent(), new TransformerWithMetadata(notifyError, videoFrameHandler)])];
            }
        });
    });
}
/**
 * Get the video track from the media stream gotten from chrome.webview.getTextureStream(streamId).
 */
function getInputVideoTrack(streamId, notifyError, videoPerformanceMonitor) {
    return videoUtils_awaiter(this, void 0, void 0, function () {
        var chrome, mediaStream, tracks, error_1, errorMsg;
        return videoUtils_generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (inServerSideRenderingEnvironment()) {
                        throw errorNotSupportedOnPlatform;
                    }
                    chrome = window['chrome'];
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportGettingTextureStream(streamId);
                    return [4 /*yield*/, chrome.webview.getTextureStream(streamId)];
                case 2:
                    mediaStream = _a.sent();
                    tracks = mediaStream.getVideoTracks();
                    if (tracks.length === 0) {
                        throw new Error("No video track in stream ".concat(streamId));
                    }
                    videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportTextureStreamAcquired();
                    return [2 /*return*/, tracks[0]];
                case 3:
                    error_1 = _a.sent();
                    errorMsg = "Failed to get video track from stream ".concat(streamId, ", error: ").concat(error_1);
                    notifyError(errorMsg);
                    throw new Error("Internal error: can't get video track from stream ".concat(streamId));
                case 4: return [2 /*return*/];
            }
        });
    });
}
/**
 * The function to create a processed video track from the original video track.
 * It reads frames from the video track and pipes them to the video frame callback to process the frames.
 * The processed frames are then enqueued to the generator.
 * The generator can be registered back to the media stream so that the host can get the processed frames.
 */
function createProcessedStreamGenerator(videoTrack, transformer) {
    if (inServerSideRenderingEnvironment()) {
        throw errorNotSupportedOnPlatform;
    }
    var MediaStreamTrackProcessor = window['MediaStreamTrackProcessor'];
    var processor = new MediaStreamTrackProcessor({ track: videoTrack, maxBufferSize: 3 });
    var source = processor.readable;
    var MediaStreamTrackGenerator = window['MediaStreamTrackGenerator'];
    var generator = new MediaStreamTrackGenerator({ kind: 'video' });
    var sink = generator.writable;
    source.pipeThrough(new TransformStream(transformer)).pipeTo(sink);
    return generator;
}
/**
 * @hidden
 * Error messages during video frame transformation.
 */
var VideoFrameTransformErrors;
(function (VideoFrameTransformErrors) {
    VideoFrameTransformErrors["TimestampIsNull"] = "timestamp of the original video frame is null";
    VideoFrameTransformErrors["UnsupportedVideoFramePixelFormat"] = "Unsupported video frame pixel format";
})(VideoFrameTransformErrors || (VideoFrameTransformErrors = {}));
var DefaultTransformer = /** @class */ (function () {
    function DefaultTransformer(notifyError, videoFrameHandler) {
        var _this = this;
        this.notifyError = notifyError;
        this.videoFrameHandler = videoFrameHandler;
        this.transform = function (originalFrame, controller) { return videoUtils_awaiter(_this, void 0, void 0, function () {
            var timestamp, frameProcessedByApp, processedFrame, error_2;
            return videoUtils_generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        timestamp = originalFrame.timestamp;
                        if (!(timestamp !== null)) return [3 /*break*/, 5];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this.videoFrameHandler({ videoFrame: originalFrame })];
                    case 2:
                        frameProcessedByApp = _a.sent();
                        processedFrame = new VideoFrame(frameProcessedByApp, {
                            // we need the timestamp to be unchanged from the oirginal frame, so we explicitly set it here.
                            timestamp: timestamp,
                        });
                        controller.enqueue(processedFrame);
                        originalFrame.close();
                        frameProcessedByApp.close();
                        return [3 /*break*/, 4];
                    case 3:
                        error_2 = _a.sent();
                        originalFrame.close();
                        this.notifyError(error_2);
                        return [3 /*break*/, 4];
                    case 4: return [3 /*break*/, 6];
                    case 5:
                        this.notifyError(VideoFrameTransformErrors.TimestampIsNull);
                        _a.label = 6;
                    case 6: return [2 /*return*/];
                }
            });
        }); };
    }
    return DefaultTransformer;
}());
/**
 * @hidden
 * Utility class to parse the header of a one-texture-input texture.
 */
var OneTextureHeader = /** @class */ (function () {
    function OneTextureHeader(headerBuffer, notifyError) {
        this.headerBuffer = headerBuffer;
        this.notifyError = notifyError;
        // Identifier for the texture layout, which is the 4-byte ASCII string "oti1" hardcoded by the host
        // (oti1 stands for "one texture input version 1")
        this.ONE_TEXTURE_INPUT_ID = 0x6f746931;
        this.INVALID_HEADER_ERROR = 'Invalid video frame header';
        this.UNSUPPORTED_LAYOUT_ERROR = 'Unsupported texture layout';
        this.headerDataView = new Uint32Array(headerBuffer);
        // headerDataView will contain the following data:
        // 0: oneTextureLayoutId
        // 1: version
        // 2: frameRowOffset
        // 3: frameFormat
        // 4: frameWidth
        // 5: frameHeight
        // 6: multiStreamHeaderRowOffset
        // 7: multiStreamCount
        if (this.headerDataView.length < 8) {
            this.notifyError(this.INVALID_HEADER_ERROR);
            throw new Error(this.INVALID_HEADER_ERROR);
        }
        // ensure the texture layout is supported
        if (this.headerDataView[0] !== this.ONE_TEXTURE_INPUT_ID) {
            this.notifyError(this.UNSUPPORTED_LAYOUT_ERROR);
            throw new Error(this.UNSUPPORTED_LAYOUT_ERROR);
        }
    }
    Object.defineProperty(OneTextureHeader.prototype, "oneTextureLayoutId", {
        get: function () {
            return this.headerDataView[0];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "version", {
        get: function () {
            return this.headerDataView[1];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "frameRowOffset", {
        get: function () {
            return this.headerDataView[2];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "frameFormat", {
        get: function () {
            return this.headerDataView[3];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "frameWidth", {
        get: function () {
            return this.headerDataView[4];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "frameHeight", {
        get: function () {
            return this.headerDataView[5];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "multiStreamHeaderRowOffset", {
        get: function () {
            return this.headerDataView[6];
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(OneTextureHeader.prototype, "multiStreamCount", {
        get: function () {
            return this.headerDataView[7];
        },
        enumerable: false,
        configurable: true
    });
    return OneTextureHeader;
}());
/**
 * @hidden
 * Utility class to parse the metadata of a one-texture-input texture.
 */
var OneTextureMetadata = /** @class */ (function () {
    function OneTextureMetadata(metadataBuffer, streamCount) {
        this.metadataMap = new Map();
        // Stream id for audio inference metadata, which is the 4-byte ASCII string "1dia" hardcoded by the host
        // (1dia stands for "audio inference data version 1")
        this.AUDIO_INFERENCE_RESULT_STREAM_ID = 0x31646961;
        var metadataDataView = new Uint32Array(metadataBuffer);
        for (var i = 0, index = 0; i < streamCount; i++) {
            var streamId = metadataDataView[index++];
            var streamDataOffset = metadataDataView[index++];
            var streamDataSize = metadataDataView[index++];
            var streamData = new Uint8Array(metadataBuffer, streamDataOffset, streamDataSize);
            this.metadataMap.set(streamId, streamData);
        }
    }
    Object.defineProperty(OneTextureMetadata.prototype, "audioInferenceResult", {
        get: function () {
            return this.metadataMap.get(this.AUDIO_INFERENCE_RESULT_STREAM_ID);
        },
        enumerable: false,
        configurable: true
    });
    return OneTextureMetadata;
}());
var TransformerWithMetadata = /** @class */ (function () {
    function TransformerWithMetadata(notifyError, videoFrameHandler) {
        var _this = this;
        this.notifyError = notifyError;
        this.videoFrameHandler = videoFrameHandler;
        this.shouldDiscardAudioInferenceResult = false;
        this.transform = function (originalFrame, controller) { return videoUtils_awaiter(_this, void 0, void 0, function () {
            var timestamp, _a, videoFrame, _b, _c, audioInferenceResult, frameProcessedByApp, processedFrame, error_3;
            return videoUtils_generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        timestamp = originalFrame.timestamp;
                        if (!(timestamp !== null)) return [3 /*break*/, 6];
                        _d.label = 1;
                    case 1:
                        _d.trys.push([1, 4, , 5]);
                        return [4 /*yield*/, this.extractVideoFrameAndMetadata(originalFrame)];
                    case 2:
                        _a = _d.sent(), videoFrame = _a.videoFrame, _b = _a.metadata, _c = _b === void 0 ? {} : _b, audioInferenceResult = _c.audioInferenceResult;
                        return [4 /*yield*/, this.videoFrameHandler({ videoFrame: videoFrame, audioInferenceResult: audioInferenceResult })];
                    case 3:
                        frameProcessedByApp = _d.sent();
                        processedFrame = new VideoFrame(frameProcessedByApp, {
                            // we need the timestamp to be unchanged from the oirginal frame, so we explicitly set it here.
                            timestamp: timestamp,
                        });
                        controller.enqueue(processedFrame);
                        videoFrame.close();
                        originalFrame.close();
                        frameProcessedByApp.close();
                        return [3 /*break*/, 5];
                    case 4:
                        error_3 = _d.sent();
                        originalFrame.close();
                        this.notifyError(error_3);
                        return [3 /*break*/, 5];
                    case 5: return [3 /*break*/, 7];
                    case 6:
                        this.notifyError(VideoFrameTransformErrors.TimestampIsNull);
                        _d.label = 7;
                    case 7: return [2 /*return*/];
                }
            });
        }); };
        /**
         * @hidden
         * Extract video frame and metadata from the given texture.
         * The given texure should be in NV12 format and the layout of the texture should be:
         * | Texture layout        |
         * | :---                  |
         * | Header                |
         * | Real video frame data |
         * | Metadata              |
         *
         * The header data is in the first two rows with the following format:
         * | oneTextureLayoutId | version | frameRowOffset | frameFormat | frameWidth | frameHeight | multiStreamHeaderRowOffset | multiStreamCount | ...   |
         * |    :---:           | :---:   | :---:          |  :---:      |  :---:     |  :---:      |  :---:                     |  :---:           | :---: |
         * | 4 bytes            | 4 bytes | 4 bytes        | 4 bytes     | 4 bytes    | 4 bytes     | 4 bytes                    | 4 bytes          | ...   |
         *
         * After header, it comes with the real video frame data.
         * At the end of the texture, it comes with the metadata. The metadata section can contain multiple types of metadata.
         * Each type of metadata is called a stream. The section is in the following format:
         * | stream1.id | stream1.dataOffset | stream1.dataSize | stream2.id | stream2.dataOffset | stream2.dataSize | ... | stream1.data | stream2.data | ... |
         * | :---:      | :---:              | :---:            |  :---:     |  :---:             |  :---:           |:---:|  :---:       | :---:        |:---:|
         * | 4 bytes    | 4 bytes            | 4 bytes          | 4 bytes    | 4 bytes            | 4 bytes          | ... | ...          | ...          | ... |
         *
         * @internal
         * Limited to Microsoft-internal use
         */
        this.extractVideoFrameAndMetadata = function (texture) { return videoUtils_awaiter(_this, void 0, void 0, function () {
            var headerRect, headerBuffer, header, metadataRect, metadataBuffer, metadata;
            return videoUtils_generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (inServerSideRenderingEnvironment()) {
                            throw errorNotSupportedOnPlatform;
                        }
                        if (texture.format !== 'NV12') {
                            this.notifyError(VideoFrameTransformErrors.UnsupportedVideoFramePixelFormat);
                            throw new Error(VideoFrameTransformErrors.UnsupportedVideoFramePixelFormat);
                        }
                        headerRect = { x: 0, y: 0, width: texture.codedWidth, height: 2 };
                        headerBuffer = new ArrayBuffer((headerRect.width * headerRect.height * 3) / 2);
                        return [4 /*yield*/, texture.copyTo(headerBuffer, { rect: headerRect })];
                    case 1:
                        _a.sent();
                        header = new OneTextureHeader(headerBuffer, this.notifyError);
                        metadataRect = {
                            x: 0,
                            y: header.multiStreamHeaderRowOffset,
                            width: texture.codedWidth,
                            height: texture.codedHeight - header.multiStreamHeaderRowOffset,
                        };
                        metadataBuffer = new ArrayBuffer((metadataRect.width * metadataRect.height * 3) / 2);
                        return [4 /*yield*/, texture.copyTo(metadataBuffer, { rect: metadataRect })];
                    case 2:
                        _a.sent();
                        metadata = new OneTextureMetadata(metadataBuffer, header.multiStreamCount);
                        return [2 /*return*/, {
                                videoFrame: new VideoFrame(texture, {
                                    timestamp: texture.timestamp,
                                    visibleRect: {
                                        x: 0,
                                        y: header.frameRowOffset,
                                        width: header.frameWidth,
                                        height: header.frameHeight,
                                    },
                                }),
                                metadata: {
                                    audioInferenceResult: this.shouldDiscardAudioInferenceResult ? undefined : metadata.audioInferenceResult,
                                },
                            }];
                }
            });
        }); };
        registerHandler('video.mediaStream.audioInferenceDiscardStatusChange', function (_a) {
            var discardAudioInferenceResult = _a.discardAudioInferenceResult;
            _this.shouldDiscardAudioInferenceResult = discardAudioInferenceResult;
        });
    }
    return TransformerWithMetadata;
}());
/**
 * @hidden
 */
function createEffectParameterChangeCallback(callback, videoPerformanceMonitor) {
    return function (effectId, effectParam) {
        videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportApplyingVideoEffect(effectId || '', effectParam);
        callback(effectId, effectParam)
            .then(function () {
            videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportVideoEffectChanged(effectId || '', effectParam);
            sendMessageToParent('video.videoEffectReadiness', [true, effectId, undefined, effectParam]);
        })
            .catch(function (reason) {
            var validReason = reason in video.EffectFailureReason ? reason : video.EffectFailureReason.InitializationFailure;
            sendMessageToParent('video.videoEffectReadiness', [false, effectId, validReason, effectParam]);
        });
    };
}

;// CONCATENATED MODULE: ./src/public/video.ts
var video_assign = (undefined && undefined.__assign) || function () {
    video_assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return video_assign.apply(this, arguments);
};
var video_awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var video_generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var __rest = (undefined && undefined.__rest) || function (s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
};








/**
 * Namespace to video extensibility of the SDK
 * @beta
 */
var video;
(function (video) {
    var videoPerformanceMonitor = inServerSideRenderingEnvironment()
        ? undefined
        : new VideoPerformanceMonitor(sendMessageToParent);
    /**
     * Video frame format enum, currently only support NV12
     * @beta
     */
    var VideoFrameFormat;
    (function (VideoFrameFormat) {
        /** Video format used for encoding and decoding YUV color data in video streaming and storage applications. */
        VideoFrameFormat["NV12"] = "NV12";
    })(VideoFrameFormat = video.VideoFrameFormat || (video.VideoFrameFormat = {}));
    /**
     * Video effect change type enum
     * @beta
     */
    var EffectChangeType;
    (function (EffectChangeType) {
        /**
         * Current video effect changed
         */
        EffectChangeType["EffectChanged"] = "EffectChanged";
        /**
         * Disable the video effect
         */
        EffectChangeType["EffectDisabled"] = "EffectDisabled";
    })(EffectChangeType = video.EffectChangeType || (video.EffectChangeType = {}));
    /**
     * Predefined failure reasons for preparing the selected video effect
     * @beta
     */
    var EffectFailureReason;
    (function (EffectFailureReason) {
        /**
         * A wrong effect id is provide.
         * Use this reason when the effect id is not found or empty, this may indicate a mismatch between the app and its manifest or a bug of the host.
         */
        EffectFailureReason["InvalidEffectId"] = "InvalidEffectId";
        /**
         * The effect can't be initialized
         */
        EffectFailureReason["InitializationFailure"] = "InitializationFailure";
    })(EffectFailureReason = video.EffectFailureReason || (video.EffectFailureReason = {}));
    /**
     * Register callbacks to process the video frames if the host supports it.
     * @beta
     * @param parameters - Callbacks and configuration to process the video frames. A host may support either {@link VideoFrameHandler} or {@link VideoBufferHandler}, but not both.
     * To ensure the video effect works on all supported hosts, the video app must provide both {@link VideoFrameHandler} and {@link VideoBufferHandler}.
     * The host will choose the appropriate callback based on the host's capability.
     *
     * @example
     * ```typescript
     * video.registerForVideoFrame({
     *   videoFrameHandler: async (videoFrameData) => {
     *     const originalFrame = videoFrameData.videoFrame as VideoFrame;
     *     try {
     *       const processedFrame = await processFrame(originalFrame);
     *       return processedFrame;
     *     } catch (e) {
     *       throw e;
     *     }
     *   },
     *   videoBufferHandler: (
     *     bufferData: VideoBufferData,
     *     notifyVideoFrameProcessed: notifyVideoFrameProcessedFunctionType,
     *     notifyError: notifyErrorFunctionType
     *     ) => {
     *       try {
     *         processFrameInplace(bufferData);
     *         notifyVideoFrameProcessed();
     *       } catch (e) {
     *         notifyError(e);
     *       }
     *     },
     *   config: {
     *     format: video.VideoPixelFormat.NV12,
     *   }
     * });
     * ```
     */
    function registerForVideoFrame(parameters) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!parameters.videoFrameHandler || !parameters.videoBufferHandler) {
            throw new Error('Both videoFrameHandler and videoBufferHandler must be provided');
        }
        registerHandler('video.setFrameProcessTimeLimit', function (timeLimitInfo) {
            return videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.setFrameProcessTimeLimit(timeLimitInfo.timeLimit);
        }, false);
        if (doesSupportMediaStream()) {
            registerForMediaStream(parameters.videoFrameHandler, parameters.config);
        }
        else if (doesSupportSharedFrame()) {
            registerForVideoBuffer(parameters.videoBufferHandler, parameters.config);
        }
        else {
            // should not happen if isSupported() is true
            throw errorNotSupportedOnPlatform;
        }
        videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.startMonitorSlowFrameProcessing();
    }
    video.registerForVideoFrame = registerForVideoFrame;
    /**
     * Video extension should call this to notify host that the current selected effect parameter changed.
     * If it's pre-meeting, host will call videoEffectCallback immediately then use the videoEffect.
     * If it's the in-meeting scenario, we will call videoEffectCallback when apply button clicked.
     * @beta
     * @param effectChangeType - the effect change type.
     * @param effectId - Newly selected effect id.
     */
    function notifySelectedVideoEffectChanged(effectChangeType, effectId) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('video.videoEffectChanged', [effectChangeType, effectId]);
    }
    video.notifySelectedVideoEffectChanged = notifySelectedVideoEffectChanged;
    /**
     * Register a callback to be notified when a new video effect is applied.
     * @beta
     * @param callback - Function to be called when new video effect is applied.
     */
    function registerForVideoEffect(callback) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('video.effectParameterChange', createEffectParameterChangeCallback(callback, videoPerformanceMonitor), false);
        sendMessageToParent('video.registerForVideoEffect');
    }
    video.registerForVideoEffect = registerForVideoEffect;
    /**
     * Sending notification to host finished the video frame processing, now host can render this video frame
     * or pass the video frame to next one in video pipeline
     * @beta
     */
    function notifyVideoFrameProcessed(timestamp) {
        sendMessageToParent('video.videoFrameProcessed', [timestamp]);
    }
    /**
     * Sending error notification to host
     * @beta
     * @param errorMessage - The error message that will be sent to the host
     */
    function notifyError(errorMessage) {
        sendMessageToParent('video.notifyError', [errorMessage]);
    }
    /**
     * Checks if video capability is supported by the host.
     * @beta
     * @returns boolean to represent whether the video capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     */
    function isSupported() {
        return (ensureInitialized(runtime) &&
            !!runtime.supports.video &&
            /** A host should support either mediaStream or sharedFrame sub-capability to support the video capability */
            (!!runtime.supports.video.mediaStream || !!runtime.supports.video.sharedFrame));
    }
    video.isSupported = isSupported;
    function registerForMediaStream(videoFrameHandler, config) {
        var _this = this;
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported() || !doesSupportMediaStream()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('video.startVideoExtensibilityVideoStream', function (mediaStreamInfo) { return video_awaiter(_this, void 0, void 0, function () {
            var streamId, monitoredVideoFrameHandler, generator;
            var _a, _b;
            return video_generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        streamId = mediaStreamInfo.streamId;
                        monitoredVideoFrameHandler = createMonitoredVideoFrameHandler(videoFrameHandler, videoPerformanceMonitor);
                        return [4 /*yield*/, processMediaStream(streamId, monitoredVideoFrameHandler, notifyError, videoPerformanceMonitor)];
                    case 1:
                        generator = _c.sent();
                        // register the video track with processed frames back to the stream:
                        !inServerSideRenderingEnvironment() && ((_b = (_a = window['chrome']) === null || _a === void 0 ? void 0 : _a.webview) === null || _b === void 0 ? void 0 : _b.registerTextureStream(streamId, generator));
                        return [2 /*return*/];
                }
            });
        }); }, false);
        sendMessageToParent('video.mediaStream.registerForVideoFrame', [config]);
    }
    function createMonitoredVideoFrameHandler(videoFrameHandler, videoPerformanceMonitor) {
        var _this = this;
        return function (videoFrameData) { return video_awaiter(_this, void 0, void 0, function () {
            var originalFrame, processedFrame;
            return video_generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        originalFrame = videoFrameData.videoFrame;
                        videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportStartFrameProcessing(originalFrame.codedWidth, originalFrame.codedHeight);
                        return [4 /*yield*/, videoFrameHandler(videoFrameData)];
                    case 1:
                        processedFrame = _a.sent();
                        videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportFrameProcessed();
                        return [2 /*return*/, processedFrame];
                }
            });
        }); };
    }
    function registerForVideoBuffer(videoBufferHandler, config) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported() || !doesSupportSharedFrame()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('video.newVideoFrame', 
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        function (videoBufferData) {
            if (videoBufferData) {
                var timestamp_1 = videoBufferData.timestamp;
                videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportStartFrameProcessing(videoBufferData.width, videoBufferData.height);
                videoBufferHandler(normalizeVideoBufferData(videoBufferData), function () {
                    videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportFrameProcessed();
                    notifyVideoFrameProcessed(timestamp_1);
                }, notifyError);
            }
        }, false);
        sendMessageToParent('video.registerForVideoFrame', [config]);
    }
    function normalizeVideoBufferData(videoBufferData) {
        if ('videoFrameBuffer' in videoBufferData) {
            return videoBufferData;
        }
        else {
            // The host may pass the VideoFrame with the old definition which has `data` instead of `videoFrameBuffer`
            var data = videoBufferData.data, newVideoBufferData = __rest(videoBufferData, ["data"]);
            return video_assign(video_assign({}, newVideoBufferData), { videoFrameBuffer: data });
        }
    }
    function doesSupportMediaStream() {
        var _a;
        return (ensureInitialized(runtime, FrameContexts.sidePanel) &&
            isTextureStreamAvailable() &&
            !!((_a = runtime.supports.video) === null || _a === void 0 ? void 0 : _a.mediaStream));
    }
    function isTextureStreamAvailable() {
        var _a, _b, _c, _d;
        return (!inServerSideRenderingEnvironment() &&
            !!(((_b = (_a = window['chrome']) === null || _a === void 0 ? void 0 : _a.webview) === null || _b === void 0 ? void 0 : _b.getTextureStream) && ((_d = (_c = window['chrome']) === null || _c === void 0 ? void 0 : _c.webview) === null || _d === void 0 ? void 0 : _d.registerTextureStream)));
    }
    function doesSupportSharedFrame() {
        var _a;
        return ensureInitialized(runtime, FrameContexts.sidePanel) && !!((_a = runtime.supports.video) === null || _a === void 0 ? void 0 : _a.sharedFrame);
    }
})(video || (video = {})); //end of video namespace

;// CONCATENATED MODULE: ./src/public/search.ts





/**
 * Allows your application to interact with the host M365 application's search box.
 * By integrating your application with the host's search box, users can search
 * your app using the same search box they use elsewhere in Teams, Outlook, or Office.
 *
 * This functionality is in Beta.
 * @beta
 */
var search;
(function (search) {
    var onChangeHandlerName = 'search.queryChange';
    var onClosedHandlerName = 'search.queryClose';
    var onExecutedHandlerName = 'search.queryExecute';
    /**
     * Allows the caller to register for various events fired by the host search experience.
     * Calling this function indicates that your application intends to plug into the host's search box and handle search events,
     * when the user is actively using your page/tab.
     *
     * The host may visually update its search box, e.g. with the name or icon of your application.
     *
     * Your application should *not* re-render inside of these callbacks, there may be a large number
     * of onChangeHandler calls if the user is typing rapidly in the search box.
     *
     * @param onClosedHandler - This handler will be called when the user exits or cancels their search.
     * Should be used to return your application to its most recent, non-search state. The value of {@link SearchQuery.searchTerm}
     * will be whatever the last query was before ending search.
     *
     * @param onExecuteHandler - The handler will be called when the user executes their
     * search (by pressing Enter for example). Should be used to display the full list of search results.
     * The value of {@link SearchQuery.searchTerm} is the complete query the user entered in the search box.
     *
     * @param onChangeHandler - This optional handler will be called when the user first starts using the
     * host's search box and as the user types their query. Can be used to put your application into a
     * word-wheeling state or to display suggestions as the user is typing.
     *
     * This handler will be called with an empty {@link SearchQuery.searchTerm} when search is beginning, and subsequently,
     * with the current contents of the search box.
     * @example
     * ``` ts
     * search.registerHandlers(
        query => {
          console.log('Update your application to handle the search experience being closed. Last query: ${query.searchTerm}');
        },
        query => {
          console.log(`Update your application to handle an executed search result: ${query.searchTerm}`);
        },
        query => {
          console.log(`Update your application with the changed search query: ${query.searchTerm}`);
        },
       );
     * ```
     *
     * @beta
     */
    function registerHandlers(onClosedHandler, onExecuteHandler, onChangeHandler) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler(onClosedHandlerName, onClosedHandler);
        registerHandler(onExecutedHandlerName, onExecuteHandler);
        if (onChangeHandler) {
            registerHandler(onChangeHandlerName, onChangeHandler);
        }
    }
    search.registerHandlers = registerHandlers;
    /**
     * Allows the caller to unregister for all events fired by the host search experience. Calling
     * this function will cause your app to stop appearing in the set of search scopes in the hosts
     *
     * @beta
     */
    function unregisterHandlers() {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        // This should let the host know to stop making the app scope show up in the search experience
        // Can also be used to clean up handlers on the host if desired
        sendMessageToParent('search.unregister');
        removeHandler(onChangeHandlerName);
        removeHandler(onClosedHandlerName);
        removeHandler(onExecutedHandlerName);
    }
    search.unregisterHandlers = unregisterHandlers;
    /**
     * Checks if search capability is supported by the host
     * @returns boolean to represent whether the search capability is supported
     *
     * @throws Error if {@link app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.search ? true : false;
    }
    search.isSupported = isSupported;
    /**
     * Clear the host M365 application's search box
     *
     * @beta
     */
    function closeSearch() {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw new Error('Not supported');
            }
            resolve(sendAndHandleStatusAndReason('search.closeSearch'));
        });
    }
    search.closeSearch = closeSearch;
})(search || (search = {}));

;// CONCATENATED MODULE: ./src/public/sharing.ts






/**
 * Namespace to open a share dialog for web content.
 * For more info, see [Share to Teams from personal app or tab](https://learn.microsoft.com/microsoftteams/platform/concepts/build-and-test/share-to-teams-from-personal-app-or-tab)
 */
var sharing;
(function (sharing) {
    /** Type of message that can be sent or received by the sharing APIs */
    sharing.SharingAPIMessages = {
        /**
         * Share web content message.
         * @internal
         */
        shareWebContent: 'sharing.shareWebContent',
    };
    function shareWebContent(shareWebContentRequest, callback) {
        // validate the given input (synchronous check)
        try {
            validateNonEmptyContent(shareWebContentRequest);
            validateTypeConsistency(shareWebContentRequest);
            validateContentForSupportedTypes(shareWebContentRequest);
        }
        catch (err) {
            //return the error via callback(v1) or rejected promise(v2)
            var wrappedFunction = function () { return Promise.reject(err); };
            return callCallbackWithSdkErrorFromPromiseAndReturnPromise(wrappedFunction, callback);
        }
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
        return callCallbackWithSdkErrorFromPromiseAndReturnPromise(shareWebContentHelper, callback, shareWebContentRequest);
    }
    sharing.shareWebContent = shareWebContent;
    function shareWebContentHelper(shareWebContentRequest) {
        return new Promise(function (resolve) {
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            resolve(sendAndHandleSdkError(sharing.SharingAPIMessages.shareWebContent, shareWebContentRequest));
        });
    }
    /**
     * Functions for validating the shareRequest input parameter
     */
    function validateNonEmptyContent(shareRequest) {
        if (!(shareRequest && shareRequest.content && shareRequest.content.length)) {
            var err = {
                errorCode: ErrorCode.INVALID_ARGUMENTS,
                message: 'Shared content is missing',
            };
            throw err;
        }
    }
    function validateTypeConsistency(shareRequest) {
        var err;
        if (shareRequest.content.some(function (item) { return !item.type; })) {
            err = {
                errorCode: ErrorCode.INVALID_ARGUMENTS,
                message: 'Shared content type cannot be undefined',
            };
            throw err;
        }
        if (shareRequest.content.some(function (item) { return item.type !== shareRequest.content[0].type; })) {
            err = {
                errorCode: ErrorCode.INVALID_ARGUMENTS,
                message: 'Shared content must be of the same type',
            };
            throw err;
        }
    }
    function validateContentForSupportedTypes(shareRequest) {
        var err;
        if (shareRequest.content[0].type === 'URL') {
            if (shareRequest.content.some(function (item) { return !item.url; })) {
                err = {
                    errorCode: ErrorCode.INVALID_ARGUMENTS,
                    message: 'URLs are required for URL content types',
                };
                throw err;
            }
        }
        else {
            err = {
                errorCode: ErrorCode.INVALID_ARGUMENTS,
                message: 'Content type is unsupported',
            };
            throw err;
        }
    }
    /**
     * Checks if the sharing capability is supported by the host
     * @returns boolean to represent whether the sharing capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.sharing ? true : false;
    }
    sharing.isSupported = isSupported;
})(sharing || (sharing = {}));

;// CONCATENATED MODULE: ./src/public/stageView.ts




/**
 * Namespace to interact with the stage view specific part of the SDK.
 *
 *  @beta
 */
var stageView;
(function (stageView) {
    /**
     * @hidden
     * Feature is under development
     *
     * Opens a stage view to display a Teams application
     * @beta
     * @param stageViewParams - The parameters to pass into the stage view.
     * @returns Promise that resolves or rejects with an error once the stage view is closed.
     */
    function open(stageViewParams) {
        return new Promise(function (resolve) {
            ensureInitialized(runtime, FrameContexts.content);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (!stageViewParams) {
                throw new Error('[stageView.open] Stage view params cannot be null');
            }
            resolve(sendAndHandleSdkError('stageView.open', stageViewParams));
        });
    }
    stageView.open = open;
    /**
     * Checks if stageView capability is supported by the host
     * @beta
     * @returns boolean to represent whether the stageView capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.stageView ? true : false;
    }
    stageView.isSupported = isSupported;
})(stageView || (stageView = {}));

;// CONCATENATED MODULE: ./src/public/webStorage.ts


/**
 * Contains functionality to allow web apps to store data in webview cache
 *
 * @beta
 */
var webStorage;
(function (webStorage) {
    /**
     * Checks if web storage gets cleared when a user logs out from host client
     *
     * @returns true is web storage gets cleared on logout and false if it does not
     *
     * @beta
     */
    function isWebStorageClearedOnUserLogOut() {
        ensureInitialized(runtime);
        return isSupported();
    }
    webStorage.isWebStorageClearedOnUserLogOut = isWebStorageClearedOnUserLogOut;
    /**
     * Checks if webStorage capability is supported by the host
     * @returns boolean to represent whether the webStorage capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.webStorage ? true : false;
    }
    webStorage.isSupported = isSupported;
})(webStorage || (webStorage = {}));

;// CONCATENATED MODULE: ./src/public/call.ts






/**
 * Used to interact with call functionality, including starting calls with other users.
 */
var call;
(function (call) {
    /** Modalities that can be associated with a call. */
    var CallModalities;
    (function (CallModalities) {
        /** Indicates that the modality is unknown or undefined. */
        CallModalities["Unknown"] = "unknown";
        /** Indicates that the call includes audio. */
        CallModalities["Audio"] = "audio";
        /** Indicates that the call includes video. */
        CallModalities["Video"] = "video";
        /** Indicates that the call includes video-based screen sharing. */
        CallModalities["VideoBasedScreenSharing"] = "videoBasedScreenSharing";
        /** Indicates that the call includes data sharing or messaging. */
        CallModalities["Data"] = "data";
    })(CallModalities = call.CallModalities || (call.CallModalities = {}));
    /**
     * Starts a call with other users
     *
     * @param startCallParams - Parameters for the call
     *
     * @throws Error if call capability is not supported
     * @throws Error if host notifies of a failed start call attempt in a legacy Teams environment
     * @returns always true if the host notifies of a successful call inititation
     */
    function startCall(startCallParams) {
        return new Promise(function (resolve) {
            var _a;
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            if (runtime.isLegacyTeams) {
                resolve(sendAndUnwrap('executeDeepLink', createTeamsDeepLinkForCall(startCallParams.targets, (_a = startCallParams.requestedModalities) === null || _a === void 0 ? void 0 : _a.includes(CallModalities.Video), startCallParams.source)).then(function (result) {
                    if (!result) {
                        throw new Error(errorCallNotStarted);
                    }
                    return result;
                }));
            }
            else {
                return sendMessageToParent('call.startCall', [startCallParams], resolve);
            }
        });
    }
    call.startCall = startCall;
    /**
     * Checks if the call capability is supported by the host
     * @returns boolean to represent whether the call capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.call ? true : false;
    }
    call.isSupported = isSupported;
})(call || (call = {}));

;// CONCATENATED MODULE: ./src/public/appInitialization.ts

/**
 * @deprecated
 * As of 2.0.0, please use {@link app} namespace instead.
 */
var appInitialization;
(function (appInitialization) {
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.Messages} instead.
     */
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    appInitialization.Messages = app.Messages;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.FailedReason} instead.
     */
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    appInitialization.FailedReason = app.FailedReason;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.ExpectedFailureReason} instead.
     */
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    appInitialization.ExpectedFailureReason = app.ExpectedFailureReason;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.notifyAppLoaded app.notifyAppLoaded(): void} instead.
     *
     * Notifies the frame that app has loaded and to hide the loading indicator if one is shown.
     */
    function notifyAppLoaded() {
        app.notifyAppLoaded();
    }
    appInitialization.notifyAppLoaded = notifyAppLoaded;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.notifySuccess app.notifySuccess(): void} instead.
     *
     * Notifies the frame that app initialization is successful and is ready for user interaction.
     */
    function notifySuccess() {
        app.notifySuccess();
    }
    appInitialization.notifySuccess = notifySuccess;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.notifyFailure app.notifyFailure(appInitializationFailedRequest: IFailedRequest): void} instead.
     *
     * Notifies the frame that app initialization has failed and to show an error page in its place.
     * @param appInitializationFailedRequest - The failure request containing the reason for why the app failed
     * during initialization as well as an optional message.
     */
    function notifyFailure(appInitializationFailedRequest) {
        app.notifyFailure(appInitializationFailedRequest);
    }
    appInitialization.notifyFailure = notifyFailure;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link app.notifyExpectedFailure app.notifyExpectedFailure(expectedFailureRequest: IExpectedFailureRequest): void} instead.
     *
     * Notifies the frame that app initialized with some expected errors.
     * @param expectedFailureRequest - The expected failure request containing the reason and an optional message
     */
    function notifyExpectedFailure(expectedFailureRequest) {
        app.notifyExpectedFailure(expectedFailureRequest);
    }
    appInitialization.notifyExpectedFailure = notifyExpectedFailure;
})(appInitialization || (appInitialization = {}));

;// CONCATENATED MODULE: ./src/public/publicAPIs.ts










/**
 * @deprecated
 * As of 2.0.0, please use {@link app.initialize app.initialize(validMessageOrigins?: string[]): Promise\<void\>} instead.
 *
 * Initializes the library. This must be called before any other SDK calls
 * but after the frame is loaded successfully.
 * @param callback - Optionally specify a callback to invoke when Teams SDK has successfully initialized
 * @param validMessageOrigins - Optionally specify a list of cross frame message origins. There must have
 * https: protocol otherwise they will be ignored. Example: https://www.example.com
 */
function initialize(callback, validMessageOrigins) {
    app.initialize(validMessageOrigins).then(function () {
        if (callback) {
            callback();
        }
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link teamsCore.enablePrintCapability teamsCore.enablePrintCapability(): void} instead.
 *
 * Enable print capability to support printing page using Ctrl+P and cmd+P
 */
function enablePrintCapability() {
    teamsCore.enablePrintCapability();
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link teamsCore.print teamsCore.print(): void} instead.
 *
 * Default print handler
 */
function print() {
    teamsCore.print();
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link app.getContext app.getContext(): Promise\<app.Context\>} instead.
 *
 * Retrieves the current context the frame is running in.
 *
 * @param callback - The callback to invoke when the {@link Context} object is retrieved.
 */
function getContext(callback) {
    ensureInitializeCalled();
    sendMessageToParent('getContext', function (context) {
        if (!context.frameContext) {
            // Fallback logic for frameContext properties
            context.frameContext = GlobalVars.frameContext;
        }
        callback(context);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link app.registerOnThemeChangeHandler app.registerOnThemeChangeHandler(handler: registerOnThemeChangeHandlerFunctionType): void} instead.
 *
 * Registers a handler for theme changes.
 * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
 *
 * @param handler - The handler to invoke when the user changes their theme.
 */
function registerOnThemeChangeHandler(handler) {
    app.registerOnThemeChangeHandler(handler);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.registerFullScreenHandler pages.registerFullScreenHandler(handler: registerFullScreenHandlerFunctionType): void} instead.
 *
 * Registers a handler for changes from or to full-screen view for a tab.
 * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
 *
 * @param handler - The handler to invoke when the user toggles full-screen view for a tab.
 */
function registerFullScreenHandler(handler) {
    registerHandlerHelper('fullScreenChange', handler, []);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.appButton.onClick pages.appButton.onClick(handler: callbackFunctionType): void} instead.
 *
 * Registers a handler for clicking the app button.
 * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
 *
 * @param handler - The handler to invoke when the personal app button is clicked in the app bar.
 */
function registerAppButtonClickHandler(handler) {
    registerHandlerHelper('appButtonClick', handler, [FrameContexts.content]);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.appButton.onHoverEnter pages.appButton.onHoverEnter(handler: callbackFunctionType): void} instead.
 *
 * Registers a handler for entering hover of the app button.
 * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
 *
 * @param handler - The handler to invoke when entering hover of the personal app button in the app bar.
 */
function registerAppButtonHoverEnterHandler(handler) {
    registerHandlerHelper('appButtonHoverEnter', handler, [FrameContexts.content]);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.appButton.onHoverLeave pages.appButton.onHoverLeave(handler: callbackFunctionType): void} instead.
 *
 * Registers a handler for exiting hover of the app button.
 * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
 * @param handler - The handler to invoke when exiting hover of the personal app button in the app bar.
 *
 */
function registerAppButtonHoverLeaveHandler(handler) {
    registerHandlerHelper('appButtonHoverLeave', handler, [FrameContexts.content]);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.backStack.registerBackButtonHandler pages.backStack.registerBackButtonHandler(handler: registerBackButtonHandlerFunctionType): void} instead.
 *
 * Registers a handler for user presses of the Team client's back button. Experiences that maintain an internal
 * navigation stack should use this handler to navigate the user back within their frame. If an app finds
 * that after running its back button handler it cannot handle the event it should call the navigateBack
 * method to ask the Teams client to handle it instead.
 *
 * @param handler - The handler to invoke when the user presses their Team client's back button.
 */
function registerBackButtonHandler(handler) {
    pages.backStack.registerBackButtonHandlerHelper(handler);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link teamsCore.registerOnLoadHandler teamsCore.registerOnLoadHandler(handler: (context: LoadContext) => void): void} instead.
 *
 * @hidden
 * Registers a handler to be called when the page has been requested to load.
 *
 * @param handler - The handler to invoke when the page is loaded.
 */
function registerOnLoadHandler(handler) {
    teamsCore.registerOnLoadHandlerHelper(handler);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link teamsCore.registerBeforeUnloadHandler teamsCore.registerBeforeUnloadHandler(handler: (readyToUnload: callbackFunctionType) => boolean): void} instead.
 *
 * @hidden
 * Registers a handler to be called before the page is unloaded.
 *
 * @param handler - The handler to invoke before the page is unloaded. If this handler returns true the page should
 * invoke the readyToUnload function provided to it once it's ready to be unloaded.
 */
function registerBeforeUnloadHandler(handler) {
    teamsCore.registerBeforeUnloadHandlerHelper(handler);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.registerFocusEnterHandler pages.registerFocusEnterHandler(handler: (navigateForward: boolean) => void): void} instead.
 *
 * @hidden
 * Registers a handler when focus needs to be passed from teams to the place of choice on app.
 *
 * @param handler - The handler to invoked by the app when they want the focus to be in the place of their choice.
 */
function registerFocusEnterHandler(handler) {
    registerHandlerHelper('focusEnter', handler, []);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.config.registerChangeConfigHandler pages.config.registerChangeConfigHandler(handler: callbackFunctionType): void} instead.
 *
 * Registers a handler for when the user reconfigurated tab.
 *
 * @param handler - The handler to invoke when the user click on Settings.
 */
function registerChangeSettingsHandler(handler) {
    registerHandlerHelper('changeSettings', handler, [FrameContexts.content]);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.tabs.getTabInstances pages.tabs.getTabInstances(tabInstanceParameters?: TabInstanceParameters): Promise\<TabInformation\>} instead.
 *
 * Allows an app to retrieve for this user tabs that are owned by this app.
 * If no TabInstanceParameters are passed, the app defaults to favorite teams and favorite channels.
 *
 * @param callback - The callback to invoke when the {@link TabInstanceParameters} object is retrieved.
 * @param tabInstanceParameters - OPTIONAL Flags that specify whether to scope call to favorite teams or channels.
 */
function getTabInstances(callback, tabInstanceParameters) {
    ensureInitialized(runtime);
    pages.tabs.getTabInstances(tabInstanceParameters).then(function (tabInfo) {
        callback(tabInfo);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.tabs.getMruTabInstances pages.tabs.getMruTabInstances(tabInstanceParameters?: TabInstanceParameters): Promise\<TabInformation\>} instead.
 *
 * Allows an app to retrieve the most recently used tabs for this user.
 *
 * @param callback - The callback to invoke when the {@link TabInformation} object is retrieved.
 * @param tabInstanceParameters - OPTIONAL Ignored, kept for future use
 */
function getMruTabInstances(callback, tabInstanceParameters) {
    ensureInitialized(runtime);
    pages.tabs.getMruTabInstances(tabInstanceParameters).then(function (tabInfo) {
        callback(tabInfo);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.shareDeepLink pages.shareDeepLink(deepLinkParameters: DeepLinkParameters): void} instead.
 *
 * Shares a deep link that a user can use to navigate back to a specific state in this page.
 *
 * @param deepLinkParameters - ID and label for the link and fallback URL.
 */
function shareDeepLink(deepLinkParameters) {
    pages.shareDeepLink({
        subPageId: deepLinkParameters.subEntityId,
        subPageLabel: deepLinkParameters.subEntityLabel,
        subPageWebUrl: deepLinkParameters.subEntityWebUrl,
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link app.openLink app.openLink(deepLink: string): Promise\<void\>} instead.
 *
 * Execute deep link API.
 *
 * @param deepLink - deep link.
 */
function executeDeepLink(deepLink, onComplete) {
    ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
    onComplete = onComplete ? onComplete : getGenericOnCompleteHandler();
    app.openLink(deepLink)
        .then(function () {
        onComplete(true);
    })
        .catch(function (err) {
        onComplete(false, err.message);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.setCurrentFrame pages.setCurrentFrame(frameInfo: FrameInfo): void} instead.
 *
 * Set the current Frame Context
 *
 * @param frameContext - FrameContext information to be set
 */
function setFrameContext(frameContext) {
    pages.setCurrentFrame(frameContext);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.initializeWithFrameContext pages.initializeWithFrameContext(frameInfo: FrameInfo, callback?: callbackFunctionType, validMessageOrigins?: string[],): void} instead.
 *
 * Initialize with FrameContext
 *
 * @param frameContext - FrameContext information to be set
 * @param callback - The optional callback to be invoked be invoked after initilizing the frame context
 * @param validMessageOrigins -  Optionally specify a list of cross frame message origins.
 * They must have https: protocol otherwise they will be ignored. Example: https:www.example.com
 */
function initializeWithFrameContext(frameContext, callback, validMessageOrigins) {
    pages.initializeWithFrameContext(frameContext, callback, validMessageOrigins);
}

;// CONCATENATED MODULE: ./src/public/navigation.ts





/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.returnFocus pages.returnFocus(navigateForward?: boolean): void} instead.
 *
 * Return focus to the main Teams app. Will focus search bar if navigating foward and app bar if navigating back.
 *
 * @param navigateForward - Determines the direction to focus in teams app.
 */
function returnFocus(navigateForward) {
    pages.returnFocus(navigateForward);
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.tabs.navigateToTab pages.tabs.navigateToTab(tabInstance: TabInstance): Promise\<void\>} instead.
 *
 * Navigates the Microsoft Teams app to the specified tab instance.
 *
 * @param tabInstance - The tab instance to navigate to.
 * @param onComplete - The callback to invoke when the action is complete.
 */
function navigateToTab(tabInstance, onComplete) {
    ensureInitialized(runtime);
    onComplete = onComplete ? onComplete : getGenericOnCompleteHandler();
    pages.tabs.navigateToTab(tabInstance)
        .then(function () {
        onComplete(true);
    })
        .catch(function (error) {
        onComplete(false, error.message);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.navigateCrossDomain pages.navigateCrossDomain(url: string): Promise\<void\>} instead.
 *
 * Navigates the frame to a new cross-domain URL. The domain of this URL must match at least one of the
 * valid domains specified in the validDomains block of the manifest; otherwise, an exception will be
 * thrown. This function needs to be used only when navigating the frame to a URL in a different domain
 * than the current one in a way that keeps the app informed of the change and allows the SDK to
 * continue working.
 *
 * @param url - The URL to navigate the frame to.
 * @param onComplete - The callback to invoke when the action is complete.
 */
function navigateCrossDomain(url, onComplete) {
    ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.settings, FrameContexts.remove, FrameContexts.task, FrameContexts.stage, FrameContexts.meetingStage);
    onComplete = onComplete ? onComplete : getGenericOnCompleteHandler();
    pages.navigateCrossDomain(url)
        .then(function () {
        onComplete(true);
    })
        .catch(function (error) {
        onComplete(false, error.message);
    });
}
/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.backStack.navigateBack pages.backStack.navigateBack(): Promise\<void\>} instead.
 *
 * Navigates back in the Teams client.
 * See registerBackButtonHandler for more information on when it's appropriate to use this method.
 *
 * @param onComplete - The callback to invoke when the action is complete.
 */
function navigateBack(onComplete) {
    ensureInitialized(runtime);
    onComplete = onComplete ? onComplete : getGenericOnCompleteHandler();
    pages.backStack.navigateBack()
        .then(function () {
        onComplete(true);
    })
        .catch(function (error) {
        onComplete(false, error.message);
    });
}

;// CONCATENATED MODULE: ./src/public/settings.ts





/**
 * @deprecated
 * As of 2.0.0, please use {@link pages.config} namespace instead.
 *
 * Namespace to interact with the settings-specific part of the SDK.
 * This object is usable only on the settings frame.
 */
var settings;
(function (settings) {
    /**
     * @deprecated
     * As of 2.0.0, please use {@link pages.config.setValidityState pages.config.setValidityState(validityState: boolean): void} instead.
     *
     * Sets the validity state for the settings.
     * The initial value is false, so the user cannot save the settings until this is called with true.
     *
     * @param validityState - Indicates whether the save or remove button is enabled for the user.
     */
    function setValidityState(validityState) {
        pages.config.setValidityState(validityState);
    }
    settings.setValidityState = setValidityState;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link pages.getConfig pages.getConfig(): Promise\<InstanceConfig\>} instead.
     *
     * Gets the settings for the current instance.
     *
     * @param callback - The callback to invoke when the {@link Settings} object is retrieved.
     */
    function getSettings(callback) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.settings, FrameContexts.remove, FrameContexts.sidePanel);
        pages.getConfig().then(function (config) {
            callback(config);
        });
    }
    settings.getSettings = getSettings;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link pages.config.setConfig pages.config.setConfig(instanceSettings: Config): Promise\<void\>} instead.
     *
     * Sets the settings for the current instance.
     * This is an asynchronous operation; calls to getSettings are not guaranteed to reflect the changed state.
     *
     * @param - Set the desired settings for this instance.
     */
    function setSettings(instanceSettings, onComplete) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.settings, FrameContexts.sidePanel);
        onComplete = onComplete ? onComplete : getGenericOnCompleteHandler();
        pages.config.setConfig(instanceSettings)
            .then(function () {
            onComplete(true);
        })
            .catch(function (error) {
            onComplete(false, error.message);
        });
    }
    settings.setSettings = setSettings;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link pages.config.registerOnSaveHandler pages.config.registerOnSaveHandler(handler: registerOnSaveHandlerFunctionType): void} instead.
     *
     * Registers a handler for when the user attempts to save the settings. This handler should be used
     * to create or update the underlying resource powering the content.
     * The object passed to the handler must be used to notify whether to proceed with the save.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the user selects the save button.
     */
    function registerOnSaveHandler(handler) {
        pages.config.registerOnSaveHandlerHelper(handler);
    }
    settings.registerOnSaveHandler = registerOnSaveHandler;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link pages.config.registerOnRemoveHandler pages.config.registerOnRemoveHandler(handler: registerOnRemoveHandlerFunctionType): void} instead.
     *
     * Registers a handler for user attempts to remove content. This handler should be used
     * to remove the underlying resource powering the content.
     * The object passed to the handler must be used to indicate whether to proceed with the removal.
     * Only one handler may be registered at a time. Subsequent registrations will override the first.
     *
     * @param handler - The handler to invoke when the user selects the remove button.
     */
    function registerOnRemoveHandler(handler) {
        pages.config.registerOnRemoveHandlerHelper(handler);
    }
    settings.registerOnRemoveHandler = registerOnRemoveHandler;
})(settings || (settings = {}));

;// CONCATENATED MODULE: ./src/public/tasks.ts
/* eslint-disable @typescript-eslint/ban-types */
var tasks_rest = (undefined && undefined.__rest) || function (s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
};






/**
 * @deprecated
 * As of 2.0.0, please use {@link dialog} namespace instead.
 *
 * Namespace to interact with the task module-specific part of the SDK.
 * This object is usable only on the content frame.
 * The tasks namespace will be deprecated. Please use dialog for future developments.
 */
var tasks;
(function (tasks) {
    /**
     * @deprecated
     * As of 2.8.0:
     * - For url-based dialogs, please use {@link dialog.url.open dialog.url.open(urlDialogInfo: UrlDialogInfo, submitHandler?: DialogSubmitHandler, messageFromChildHandler?: PostMessageChannel): void} .
     * - For url-based dialogs with bot interaction, please use {@link dialog.url.bot.open dialog.url.bot.open(botUrlDialogInfo: BotUrlDialogInfo, submitHandler?: DialogSubmitHandler, messageFromChildHandler?: PostMessageChannel): void}
     * - For Adaptive Card-based dialogs:
     *   - In Teams, please continue to use this function until the new functions in {@link dialog.adaptiveCard} have been fully implemented. You can tell at runtime when they are implemented in Teams by checking
     *     the return value of the {@link dialog.adaptiveCard.isSupported} function. This documentation line will also be removed once they have been fully implemented in Teams.
     *   - In all other hosts, please use {@link dialog.adaptiveCard.open dialog.adaptiveCard.open(cardDialogInfo: CardDialogInfo, submitHandler?: DialogSubmitHandler, messageFromChildHandler?: PostMessageChannel): void}
     *
     * Allows an app to open the task module.
     *
     * @param taskInfo - An object containing the parameters of the task module
     * @param submitHandler - Handler to call when the task module is completed
     */
    function startTask(taskInfo, submitHandler) {
        var dialogSubmitHandler = submitHandler
            ? /* eslint-disable-next-line strict-null-checks/all */ /* fix tracked by 5730662 */
                function (sdkResponse) { return submitHandler(sdkResponse.err, sdkResponse.result); }
            : undefined;
        if (taskInfo.card === undefined && taskInfo.url === undefined) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
            sendMessageToParent('tasks.startTask', [taskInfo], submitHandler);
        }
        else if (taskInfo.card) {
            ensureInitialized(runtime, FrameContexts.content, FrameContexts.sidePanel, FrameContexts.meetingStage);
            sendMessageToParent('tasks.startTask', [taskInfo], submitHandler);
        }
        else if (taskInfo.completionBotId !== undefined) {
            dialog.url.bot.open(getBotUrlDialogInfoFromTaskInfo(taskInfo), dialogSubmitHandler);
        }
        else {
            dialog.url.open(getUrlDialogInfoFromTaskInfo(taskInfo), dialogSubmitHandler);
        }
        return new ChildAppWindow();
    }
    tasks.startTask = startTask;
    /**
     * @deprecated
     * As of 2.0.0, please use {@link dialog.update.resize dialog.update.resize(dimensions: DialogSize): void} instead.
     *
     * Update height/width task info properties.
     *
     * @param taskInfo - An object containing width and height properties
     */
    function updateTask(taskInfo) {
        taskInfo = getDefaultSizeIfNotProvided(taskInfo);
        // eslint-disable-next-line @typescript-eslint/no-unused-vars
        var width = taskInfo.width, height = taskInfo.height, extra = tasks_rest(taskInfo, ["width", "height"]);
        if (Object.keys(extra).length) {
            throw new Error('resize requires a TaskInfo argument containing only width and height');
        }
        dialog.update.resize(taskInfo);
    }
    tasks.updateTask = updateTask;
    /**
     * @deprecated
     * As of 2.8.0, please use {@link dialog.url.submit} instead.
     *
     * Submit the task module.
     *
     * @param result - Contains the result to be sent to the bot or the app. Typically a JSON object or a serialized version of it
     * @param appIds - Valid application(s) that can receive the result of the submitted dialogs. Specifying this parameter helps prevent malicious apps from retrieving the dialog result. Multiple app IDs can be specified because a web app from a single underlying domain can power multiple apps across different environments and branding schemes.
     */
    function submitTask(result, appIds) {
        dialog.url.submit(result, appIds);
    }
    tasks.submitTask = submitTask;
    /**
     * Converts {@link TaskInfo} to {@link UrlDialogInfo}
     * @param taskInfo - TaskInfo object to convert
     * @returns - Converted UrlDialogInfo object
     */
    function getUrlDialogInfoFromTaskInfo(taskInfo) {
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        var urldialogInfo = {
            url: taskInfo.url,
            size: {
                height: taskInfo.height ? taskInfo.height : TaskModuleDimension.Small,
                width: taskInfo.width ? taskInfo.width : TaskModuleDimension.Small,
            },
            title: taskInfo.title,
            fallbackUrl: taskInfo.fallbackUrl,
        };
        return urldialogInfo;
    }
    /**
     * Converts {@link TaskInfo} to {@link BotUrlDialogInfo}
     * @param taskInfo - TaskInfo object to convert
     * @returns - converted BotUrlDialogInfo object
     */
    function getBotUrlDialogInfoFromTaskInfo(taskInfo) {
        /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
        var botUrldialogInfo = {
            url: taskInfo.url,
            size: {
                height: taskInfo.height ? taskInfo.height : TaskModuleDimension.Small,
                width: taskInfo.width ? taskInfo.width : TaskModuleDimension.Small,
            },
            title: taskInfo.title,
            fallbackUrl: taskInfo.fallbackUrl,
            completionBotId: taskInfo.completionBotId,
        };
        return botUrldialogInfo;
    }
    /**
     * Sets the height and width of the {@link TaskInfo} object to the original height and width, if initially specified,
     * otherwise uses the height and width values corresponding to {@link DialogDimension | TaskModuleDimension.Small}
     * @param taskInfo TaskInfo object from which to extract size info, if specified
     * @returns TaskInfo with height and width specified
     */
    function getDefaultSizeIfNotProvided(taskInfo) {
        taskInfo.height = taskInfo.height ? taskInfo.height : TaskModuleDimension.Small;
        taskInfo.width = taskInfo.width ? taskInfo.width : TaskModuleDimension.Small;
        return taskInfo;
    }
    tasks.getDefaultSizeIfNotProvided = getDefaultSizeIfNotProvided;
})(tasks || (tasks = {}));

;// CONCATENATED MODULE: ./src/public/liveShareHost.ts




/**
 * APIs involving Live Share, a framework for building real-time collaborative apps.
 * For more information, visit https://aka.ms/teamsliveshare
 *
 * @see LiveShareHost
 */
var liveShare;
(function (liveShare) {
    /**
     * @hidden
     * The meeting roles of a user.
     * Used in Live Share for its role verification feature.
     * For more information, visit https://learn.microsoft.com/microsoftteams/platform/apps-in-teams-meetings/teams-live-share-capabilities?tabs=javascript#role-verification-for-live-data-structures
     */
    var UserMeetingRole;
    (function (UserMeetingRole) {
        /**
         * Guest role.
         */
        UserMeetingRole["guest"] = "Guest";
        /**
         * Attendee role.
         */
        UserMeetingRole["attendee"] = "Attendee";
        /**
         * Presenter role.
         */
        UserMeetingRole["presenter"] = "Presenter";
        /**
         * Organizer role.
         */
        UserMeetingRole["organizer"] = "Organizer";
    })(UserMeetingRole = liveShare.UserMeetingRole || (liveShare.UserMeetingRole = {}));
    /**
     * @hidden
     * State of the current Live Share session's Fluid container.
     * This is used internally by the `LiveShareClient` when joining a Live Share session.
     */
    var ContainerState;
    (function (ContainerState) {
        /**
         * The call to `LiveShareHost.setContainerId()` successfully created the container mapping
         * for the current Live Share session.
         */
        ContainerState["added"] = "Added";
        /**
         * A container mapping for the current Live Share session already exists.
         * This indicates to Live Share that a new container does not need be created.
         */
        ContainerState["alreadyExists"] = "AlreadyExists";
        /**
         * The call to `LiveShareHost.setContainerId()` failed to create the container mapping.
         * This happens when another client has already set the container ID for the session.
         */
        ContainerState["conflict"] = "Conflict";
        /**
         * A container mapping for the current Live Share session does not yet exist.
         * This indicates to Live Share that a new container should be created.
         */
        ContainerState["notFound"] = "NotFound";
    })(ContainerState = liveShare.ContainerState || (liveShare.ContainerState = {}));
    /**
     * Checks if the interactive capability is supported by the host
     * @returns boolean to represent whether the interactive capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     */
    function isSupported() {
        return ensureInitialized(runtime, FrameContexts.meetingStage, FrameContexts.sidePanel) &&
            runtime.supports.interactive
            ? true
            : false;
    }
    liveShare.isSupported = isSupported;
})(liveShare || (liveShare = {}));
/**
 * Live Share host implementation for connecting to real-time collaborative sessions.
 * Designed for use with the `LiveShareClient` class in the `@microsoft/live-share` package.
 * Learn more at https://aka.ms/teamsliveshare
 *
 * @remarks
 * The `LiveShareClient` class from Live Share uses the hidden API's to join/manage the session.
 * To create a new `LiveShareHost` instance use the static `LiveShareHost.create()` function.
 */
var LiveShareHost = /** @class */ (function () {
    function LiveShareHost() {
    }
    /**
     * @hidden
     * Returns the Fluid Tenant connection info for user's current context.
     */
    LiveShareHost.prototype.getFluidTenantInfo = function () {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.getFluidTenantInfo'));
        });
    };
    /**
     * @hidden
     * Returns the fluid access token for mapped container Id.
     *
     * @param containerId Fluid's container Id for the request. Undefined for new containers.
     * @returns token for connecting to Fluid's session.
     */
    LiveShareHost.prototype.getFluidToken = function (containerId) {
        ensureSupported();
        return new Promise(function (resolve) {
            // eslint-disable-next-line strict-null-checks/all
            resolve(sendAndHandleSdkError('interactive.getFluidToken', containerId));
        });
    };
    /**
     * @hidden
     * Returns the ID of the fluid container associated with the user's current context.
     */
    LiveShareHost.prototype.getFluidContainerId = function () {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.getFluidContainerId'));
        });
    };
    /**
     * @hidden
     * Sets the ID of the fluid container associated with the current context.
     *
     * @remarks
     * If this returns false, the client should delete the container they created and then call
     * `getFluidContainerId()` to get the ID of the container being used.
     * @param containerId ID of the fluid container the client created.
     * @returns A data structure with a `containerState` indicating the success or failure of the request.
     */
    LiveShareHost.prototype.setFluidContainerId = function (containerId) {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.setFluidContainerId', containerId));
        });
    };
    /**
     * @hidden
     * Returns the shared clock server's current time.
     */
    LiveShareHost.prototype.getNtpTime = function () {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.getNtpTime'));
        });
    };
    /**
     * @hidden
     * Associates the fluid client ID with a set of user roles.
     *
     * @param clientId The ID for the current user's Fluid client. Changes on reconnects.
     * @returns The roles for the current user.
     */
    LiveShareHost.prototype.registerClientId = function (clientId) {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.registerClientId', clientId));
        });
    };
    /**
     * @hidden
     * Returns the roles associated with a client ID.
     *
     * @param clientId The Client ID the message was received from.
     * @returns The roles for a given client. Returns `undefined` if the client ID hasn't been registered yet.
     */
    LiveShareHost.prototype.getClientRoles = function (clientId) {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.getClientRoles', clientId));
        });
    };
    /**
     * @hidden
     * Returns the `IClientInfo` associated with a client ID.
     *
     * @param clientId The Client ID the message was received from.
     * @returns The info for a given client. Returns `undefined` if the client ID hasn't been registered yet.
     */
    LiveShareHost.prototype.getClientInfo = function (clientId) {
        ensureSupported();
        return new Promise(function (resolve) {
            resolve(sendAndHandleSdkError('interactive.getClientInfo', clientId));
        });
    };
    /**
     * Factories a new `LiveShareHost` instance for use with the `LiveShareClient` class
     * in the `@microsoft/live-share` package.
     *
     * @remarks
     * `app.initialize()` must first be called before using this API.
     * This API can only be called from `meetingStage` or `sidePanel` contexts.
     */
    LiveShareHost.create = function () {
        ensureSupported();
        return new LiveShareHost();
    };
    return LiveShareHost;
}());

function ensureSupported() {
    if (!liveShare.isSupported()) {
        throw new Error('LiveShareHost Not supported');
    }
}

;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/regex.js
/* harmony default export */ const regex = (/^(?:[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}|00000000-0000-0000-0000-000000000000)$/i);
;// CONCATENATED MODULE: ../../node_modules/.pnpm/uuid@9.0.0/node_modules/uuid/dist/esm-browser/validate.js


function validate_validate(uuid) {
  return typeof uuid === 'string' && regex.test(uuid);
}

/* harmony default export */ const esm_browser_validate = (validate_validate);
;// CONCATENATED MODULE: ./src/internal/marketplaceUtils.ts
var marketplaceUtils_assign = (undefined && undefined.__assign) || function () {
    marketplaceUtils_assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return marketplaceUtils_assign.apply(this, arguments);
};
var marketplaceUtils_rest = (undefined && undefined.__rest) || function (s, e) {
    var t = {};
    for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p) && e.indexOf(p) < 0)
        t[p] = s[p];
    if (s != null && typeof Object.getOwnPropertySymbols === "function")
        for (var i = 0, p = Object.getOwnPropertySymbols(s); i < p.length; i++) {
            if (e.indexOf(p[i]) < 0 && Object.prototype.propertyIsEnumerable.call(s, p[i]))
                t[p[i]] = s[p[i]];
        }
    return t;
};
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */
/* eslint-disable @typescript-eslint/no-explicit-any */


/**
 * @hidden
 * deserialize the cart data:
 * - convert url properties from string to URL
 * @param cartItems The cart items
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function deserializeCart(cartData) {
    try {
        cartData.cartItems = deserializeCartItems(cartData.cartItems);
        return cartData;
    }
    catch (e) {
        throw new Error('Error deserializing cart');
    }
}
/**
 * @hidden
 * deserialize the cart items:
 * - convert url properties from string to URL
 * @param cartItems The cart items
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function deserializeCartItems(cartItemsData) {
    return cartItemsData.map(function (cartItem) {
        if (cartItem.imageURL) {
            var url = new URL(cartItem.imageURL);
            cartItem.imageURL = url;
        }
        if (cartItem.accessories) {
            cartItem.accessories = deserializeCartItems(cartItem.accessories);
        }
        return cartItem;
    });
}
/**
 * @hidden
 * serialize the cart items:
 * - make URL properties to string
 * @param cartItems The cart items
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var serializeCartItems = function (cartItems) {
    try {
        return cartItems.map(function (cartItem) {
            var imageURL = cartItem.imageURL, accessories = cartItem.accessories, rest = marketplaceUtils_rest(cartItem, ["imageURL", "accessories"]);
            var cartItemsData = marketplaceUtils_assign({}, rest);
            if (imageURL) {
                cartItemsData.imageURL = imageURL.href;
            }
            if (accessories) {
                cartItemsData.accessories = serializeCartItems(accessories);
            }
            return cartItemsData;
        });
    }
    catch (e) {
        throw new Error('Error serializing cart items');
    }
};
/**
 * @hidden
 * Validate the cart item properties are valid
 * @param cartItems The cart items
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateCartItems(cartItems) {
    if (!Array.isArray(cartItems) || cartItems.length === 0) {
        throw new Error('cartItems must be a non-empty array');
    }
    for (var _i = 0, cartItems_1 = cartItems; _i < cartItems_1.length; _i++) {
        var cartItem = cartItems_1[_i];
        validateBasicCartItem(cartItem);
        validateAccessoryItems(cartItem.accessories);
    }
}
/**
 * @hidden
 * Validate accessories
 * @param accessoryItems The accessories to be validated
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateAccessoryItems(accessoryItems) {
    if (accessoryItems === null || accessoryItems === undefined) {
        return;
    }
    if (!Array.isArray(accessoryItems) || accessoryItems.length === 0) {
        throw new Error('CartItem.accessories must be a non-empty array');
    }
    for (var _i = 0, accessoryItems_1 = accessoryItems; _i < accessoryItems_1.length; _i++) {
        var accessoryItem = accessoryItems_1[_i];
        if (accessoryItem['accessories']) {
            throw new Error('Item in CartItem.accessories cannot have accessories');
        }
        validateBasicCartItem(accessoryItem);
    }
}
/**
 * @hidden
 * Validate the basic cart item properties are valid
 * @param basicCartItem The basic cart item
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateBasicCartItem(basicCartItem) {
    if (!basicCartItem.id) {
        throw new Error('cartItem.id must not be empty');
    }
    if (!basicCartItem.name) {
        throw new Error('cartItem.name must not be empty');
    }
    validatePrice(basicCartItem.price);
    validateQuantity(basicCartItem.quantity);
}
/**
 * @hidden
 * Validate the id is valid
 * @param id A uuid string
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateUuid(id) {
    if (id === undefined || id === null) {
        return;
    }
    if (!id) {
        throw new Error('id must not be empty');
    }
    if (esm_browser_validate(id) === false) {
        throw new Error('id must be a valid UUID');
    }
}
/**
 * @hidden
 * Validate the cart item properties are valid
 * @param price The price to be validated
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validatePrice(price) {
    if (typeof price !== 'number' || price < 0) {
        throw new Error("price ".concat(price, " must be a number not less than 0"));
    }
    if (parseFloat(price.toFixed(3)) !== price) {
        throw new Error("price ".concat(price, " must have at most 3 decimal places"));
    }
}
/**
 * @hidden
 * Validate quantity
 * @param quantity The quantity to be validated
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateQuantity(quantity) {
    if (typeof quantity !== 'number' || quantity <= 0 || parseInt(quantity.toString()) !== quantity) {
        throw new Error("quantity ".concat(quantity, " must be an integer greater than 0"));
    }
}
/**
 * @hidden
 * Validate cart status
 * @param cartStatus The cartStatus to be validated
 *
 * @internal
 * Limited to Microsoft-internal use
 */
function validateCartStatus(cartStatus) {
    if (!Object.values(marketplace.CartStatus).includes(cartStatus)) {
        throw new Error("cartStatus ".concat(cartStatus, " is not valid"));
    }
}

;// CONCATENATED MODULE: ./src/public/marketplace.ts
var marketplace_assign = (undefined && undefined.__assign) || function () {
    marketplace_assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return marketplace_assign.apply(this, arguments);
};





/**
 * @hidden
 * Namespace for an app to support a checkout flow by interacting with the marketplace cart in the host.
 * @beta
 */
var marketplace;
(function (marketplace) {
    /**
     * @hidden
     * the version of the current cart interface
     * which is forced to send to the host in the calls.
     * @internal
     * Limited to Microsoft-internal use
     * @beta
     */
    marketplace.cartVersion = {
        /**
         * @hidden
         * Represents the major version of the current cart interface,
         * it is increased when incompatible interface update happens.
         */
        majorVersion: 1,
        /**
         * @hidden
         * The minor version of the current cart interface, which is compatible
         * with the previous minor version in the same major version.
         */
        minorVersion: 1,
    };
    /**
     * @hidden
     * Represents the persona creating the cart.
     * @beta
     */
    var Intent;
    (function (Intent) {
        /**
         * @hidden
         * The cart is created by admin of an organization in Teams Admin Center.
         */
        Intent["TACAdminUser"] = "TACAdminUser";
        /**
         * @hidden
         * The cart is created by admin of an organization in Teams.
         */
        Intent["TeamsAdminUser"] = "TeamsAdminUser";
        /**
         * @hidden
         * The cart is created by end user of an organization in Teams.
         */
        Intent["TeamsEndUser"] = "TeamsEndUser";
    })(Intent = marketplace.Intent || (marketplace.Intent = {}));
    /**
     * @hidden
     * Represents the status of the cart.
     * @beta
     */
    var CartStatus;
    (function (CartStatus) {
        /**
         * @hidden
         * Cart is created but not checked out yet.
         */
        CartStatus["Open"] = "Open";
        /**
         * @hidden
         * Cart is checked out but not completed yet.
         */
        CartStatus["Processing"] = "Processing";
        /**
         * @hidden
         * Indicate checking out is completed and the host should
         * return a new cart in the next getCart call.
         */
        CartStatus["Processed"] = "Processed";
        /**
         * @hidden
         * Indicate checking out process is manually cancelled by the user
         */
        CartStatus["Closed"] = "Closed";
        /**
         * @hidden
         * Indicate checking out is failed and the host should
         * return a new cart in the next getCart call.
         */
        CartStatus["Error"] = "Error";
    })(CartStatus = marketplace.CartStatus || (marketplace.CartStatus = {}));
    /**
     * @hidden
     * Get the cart object owned by the host to checkout.
     * @returns A promise of the cart object in the cartVersion.
     * @beta
     */
    function getCart() {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        return sendAndHandleSdkError('marketplace.getCart', marketplace.cartVersion).then(deserializeCart);
    }
    marketplace.getCart = getCart;
    /**
     * @hidden
     * Add or update cart items in the cart owned by the host.
     * @param addOrUpdateCartItemsParams Represents the parameters to update the cart items.
     * @returns A promise of the updated cart object in the cartVersion.
     * @beta
     */
    function addOrUpdateCartItems(addOrUpdateCartItemsParams) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!addOrUpdateCartItemsParams) {
            throw new Error('addOrUpdateCartItemsParams must be provided');
        }
        validateUuid(addOrUpdateCartItemsParams === null || addOrUpdateCartItemsParams === void 0 ? void 0 : addOrUpdateCartItemsParams.cartId);
        validateCartItems(addOrUpdateCartItemsParams === null || addOrUpdateCartItemsParams === void 0 ? void 0 : addOrUpdateCartItemsParams.cartItems);
        return sendAndHandleSdkError('marketplace.addOrUpdateCartItems', marketplace_assign(marketplace_assign({}, addOrUpdateCartItemsParams), { cartItems: serializeCartItems(addOrUpdateCartItemsParams.cartItems), cartVersion: marketplace.cartVersion })).then(deserializeCart);
    }
    marketplace.addOrUpdateCartItems = addOrUpdateCartItems;
    /**
     * @hidden
     * Remove cart items from the cart owned by the host.
     * @param removeCartItemsParams The parameters to remove the cart items.
     * @returns A promise of the updated cart object in the cartVersion.
     * @beta
     */
    function removeCartItems(removeCartItemsParams) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!removeCartItemsParams) {
            throw new Error('removeCartItemsParams must be provided');
        }
        validateUuid(removeCartItemsParams === null || removeCartItemsParams === void 0 ? void 0 : removeCartItemsParams.cartId);
        if (!Array.isArray(removeCartItemsParams === null || removeCartItemsParams === void 0 ? void 0 : removeCartItemsParams.cartItemIds) || (removeCartItemsParams === null || removeCartItemsParams === void 0 ? void 0 : removeCartItemsParams.cartItemIds.length) === 0) {
            throw new Error('cartItemIds must be a non-empty array');
        }
        return sendAndHandleSdkError('marketplace.removeCartItems', marketplace_assign(marketplace_assign({}, removeCartItemsParams), { cartVersion: marketplace.cartVersion })).then(deserializeCart);
    }
    marketplace.removeCartItems = removeCartItems;
    /**
     * @hidden
     * Update cart status in the cart owned by the host.
     * @param updateCartStatusParams The parameters to update the cart status.
     * @returns A promise of the updated cart object in the cartVersion.
     * @beta
     */
    function updateCartStatus(updateCartStatusParams) {
        ensureInitialized(runtime, FrameContexts.content, FrameContexts.task);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!updateCartStatusParams) {
            throw new Error('updateCartStatusParams must be provided');
        }
        validateUuid(updateCartStatusParams === null || updateCartStatusParams === void 0 ? void 0 : updateCartStatusParams.cartId);
        validateCartStatus(updateCartStatusParams === null || updateCartStatusParams === void 0 ? void 0 : updateCartStatusParams.cartStatus);
        return sendAndHandleSdkError('marketplace.updateCartStatus', marketplace_assign(marketplace_assign({}, updateCartStatusParams), { cartVersion: marketplace.cartVersion })).then(deserializeCart);
    }
    marketplace.updateCartStatus = updateCartStatus;
    /**
     * @hidden
     * Checks if the marketplace capability is supported by the host.
     * @returns Boolean to represent whether the marketplace capability is supported.
     * @throws Error if {@linkcode app.initialize} has not successfully completed.
     * @beta
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.marketplace ? true : false;
    }
    marketplace.isSupported = isSupported;
})(marketplace || (marketplace = {}));

;// CONCATENATED MODULE: ./src/public/index.ts







































;// CONCATENATED MODULE: ./src/private/files.ts





/**
 * @hidden
 *
 * Namespace to interact with the files specific part of the SDK.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var files;
(function (files_1) {
    /**
     * @hidden
     *
     * Cloud storage providers registered with Microsoft Teams
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var CloudStorageProvider;
    (function (CloudStorageProvider) {
        CloudStorageProvider["Dropbox"] = "DROPBOX";
        CloudStorageProvider["Box"] = "BOX";
        CloudStorageProvider["Sharefile"] = "SHAREFILE";
        CloudStorageProvider["GoogleDrive"] = "GOOGLEDRIVE";
        CloudStorageProvider["Egnyte"] = "EGNYTE";
        CloudStorageProvider["SharePoint"] = "SharePoint";
    })(CloudStorageProvider = files_1.CloudStorageProvider || (files_1.CloudStorageProvider = {}));
    /**
     * @hidden
     *
     * Cloud storage provider type enums
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var CloudStorageProviderType;
    (function (CloudStorageProviderType) {
        CloudStorageProviderType[CloudStorageProviderType["Sharepoint"] = 0] = "Sharepoint";
        CloudStorageProviderType[CloudStorageProviderType["WopiIntegration"] = 1] = "WopiIntegration";
        CloudStorageProviderType[CloudStorageProviderType["Google"] = 2] = "Google";
        CloudStorageProviderType[CloudStorageProviderType["OneDrive"] = 3] = "OneDrive";
        CloudStorageProviderType[CloudStorageProviderType["Recent"] = 4] = "Recent";
        CloudStorageProviderType[CloudStorageProviderType["Aggregate"] = 5] = "Aggregate";
        CloudStorageProviderType[CloudStorageProviderType["FileSystem"] = 6] = "FileSystem";
        CloudStorageProviderType[CloudStorageProviderType["Search"] = 7] = "Search";
        CloudStorageProviderType[CloudStorageProviderType["AllFiles"] = 8] = "AllFiles";
        CloudStorageProviderType[CloudStorageProviderType["SharedWithMe"] = 9] = "SharedWithMe";
    })(CloudStorageProviderType = files_1.CloudStorageProviderType || (files_1.CloudStorageProviderType = {}));
    /**
     * @hidden
     *
     * Special Document Library enum
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var SpecialDocumentLibraryType;
    (function (SpecialDocumentLibraryType) {
        SpecialDocumentLibraryType["ClassMaterials"] = "classMaterials";
    })(SpecialDocumentLibraryType = files_1.SpecialDocumentLibraryType || (files_1.SpecialDocumentLibraryType = {}));
    /**
     * @hidden
     *
     * Document Library Access enum
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var DocumentLibraryAccessType;
    (function (DocumentLibraryAccessType) {
        DocumentLibraryAccessType["Readonly"] = "readonly";
    })(DocumentLibraryAccessType = files_1.DocumentLibraryAccessType || (files_1.DocumentLibraryAccessType = {}));
    /**
     * @hidden
     *
     * Download status enum
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var FileDownloadStatus;
    (function (FileDownloadStatus) {
        FileDownloadStatus["Downloaded"] = "Downloaded";
        FileDownloadStatus["Downloading"] = "Downloading";
        FileDownloadStatus["Failed"] = "Failed";
    })(FileDownloadStatus = files_1.FileDownloadStatus || (files_1.FileDownloadStatus = {}));
    /**
     * @hidden
     * Hide from docs
     *
     * Actions specific to 3P cloud storage provider file and / or account
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var CloudStorageProviderFileAction;
    (function (CloudStorageProviderFileAction) {
        CloudStorageProviderFileAction["Download"] = "DOWNLOAD";
        CloudStorageProviderFileAction["Upload"] = "UPLOAD";
        CloudStorageProviderFileAction["Delete"] = "DELETE";
    })(CloudStorageProviderFileAction = files_1.CloudStorageProviderFileAction || (files_1.CloudStorageProviderFileAction = {}));
    /**
     * @hidden
     * Hide from docs
     *
     * Gets a list of cloud storage folders added to the channel
     * @param channelId - ID of the channel whose cloud storage folders should be retrieved
     * @param callback - Callback that will be triggered post folders load
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getCloudStorageFolders(channelId, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!channelId || channelId.length === 0) {
            throw new Error('[files.getCloudStorageFolders] channelId name cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[files.getCloudStorageFolders] Callback cannot be null');
        }
        sendMessageToParent('files.getCloudStorageFolders', [channelId], callback);
    }
    files_1.getCloudStorageFolders = getCloudStorageFolders;
    /**
     * @hidden
     * Hide from docs
     * ------
     * Initiates the add cloud storage folder flow
     *
     * @param channelId - ID of the channel to add cloud storage folder
     * @param callback - Callback that will be triggered post add folder flow is compelete
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function addCloudStorageFolder(channelId, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!channelId || channelId.length === 0) {
            throw new Error('[files.addCloudStorageFolder] channelId name cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[files.addCloudStorageFolder] Callback cannot be null');
        }
        sendMessageToParent('files.addCloudStorageFolder', [channelId], callback);
    }
    files_1.addCloudStorageFolder = addCloudStorageFolder;
    /**
     * @hidden
     * Hide from docs
     * ------
     *
     * Deletes a cloud storage folder from channel
     *
     * @param channelId - ID of the channel where folder is to be deleted
     * @param folderToDelete - cloud storage folder to be deleted
     * @param callback - Callback that will be triggered post delete
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function deleteCloudStorageFolder(channelId, folderToDelete, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!channelId) {
            throw new Error('[files.deleteCloudStorageFolder] channelId name cannot be null or empty');
        }
        if (!folderToDelete) {
            throw new Error('[files.deleteCloudStorageFolder] folderToDelete cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[files.deleteCloudStorageFolder] Callback cannot be null');
        }
        sendMessageToParent('files.deleteCloudStorageFolder', [channelId, folderToDelete], callback);
    }
    files_1.deleteCloudStorageFolder = deleteCloudStorageFolder;
    /**
     * @hidden
     * Hide from docs
     * ------
     *
     * Fetches the contents of a Cloud storage folder (CloudStorageFolder) / sub directory
     *
     * @param folder - Cloud storage folder (CloudStorageFolder) / sub directory (CloudStorageFolderItem)
     * @param providerCode - Code of the cloud storage folder provider
     * @param callback - Callback that will be triggered post contents are loaded
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getCloudStorageFolderContents(folder, providerCode, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!folder || !providerCode) {
            throw new Error('[files.getCloudStorageFolderContents] folder/providerCode name cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[files.getCloudStorageFolderContents] Callback cannot be null');
        }
        if ('isSubdirectory' in folder && !folder.isSubdirectory) {
            throw new Error('[files.getCloudStorageFolderContents] provided folder is not a subDirectory');
        }
        sendMessageToParent('files.getCloudStorageFolderContents', [folder, providerCode], callback);
    }
    files_1.getCloudStorageFolderContents = getCloudStorageFolderContents;
    /**
     * @hidden
     * Hide from docs
     * ------
     *
     * Open a cloud storage file in Teams
     *
     * @param file - cloud storage file that should be opened
     * @param providerCode - Code of the cloud storage folder provider
     * @param fileOpenPreference - Whether file should be opened in web/inline
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function openCloudStorageFile(file, providerCode, fileOpenPreference) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!file || !providerCode) {
            throw new Error('[files.openCloudStorageFile] file/providerCode cannot be null or empty');
        }
        if (file.isSubdirectory) {
            throw new Error('[files.openCloudStorageFile] provided file is a subDirectory');
        }
        sendMessageToParent('files.openCloudStorageFile', [file, providerCode, fileOpenPreference]);
    }
    files_1.openCloudStorageFile = openCloudStorageFile;
    /**
     * @hidden
     * Allow 1st party apps to call this function to get the external
     * third party cloud storage accounts that the tenant supports
     * @param excludeAddedProviders: return a list of support third party
     * cloud storages that hasn't been added yet.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getExternalProviders(excludeAddedProviders, callback) {
        if (excludeAddedProviders === void 0) { excludeAddedProviders = false; }
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw new Error('[files.getExternalProviders] Callback cannot be null');
        }
        sendMessageToParent('files.getExternalProviders', [excludeAddedProviders], callback);
    }
    files_1.getExternalProviders = getExternalProviders;
    /**
     * @hidden
     * Allow 1st party apps to call this function to move files
     * among SharePoint and third party cloud storages.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function copyMoveFiles(selectedFiles, providerCode, destinationFolder, destinationProviderCode, isMove, callback) {
        if (isMove === void 0) { isMove = false; }
        ensureInitialized(runtime, FrameContexts.content);
        if (!selectedFiles || selectedFiles.length === 0) {
            throw new Error('[files.copyMoveFiles] selectedFiles cannot be null or empty');
        }
        if (!providerCode) {
            throw new Error('[files.copyMoveFiles] providerCode cannot be null or empty');
        }
        if (!destinationFolder) {
            throw new Error('[files.copyMoveFiles] destinationFolder cannot be null or empty');
        }
        if (!destinationProviderCode) {
            throw new Error('[files.copyMoveFiles] destinationProviderCode cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[files.copyMoveFiles] callback cannot be null');
        }
        sendMessageToParent('files.copyMoveFiles', [selectedFiles, providerCode, destinationFolder, destinationProviderCode, isMove], callback);
    }
    files_1.copyMoveFiles = copyMoveFiles;
    /**
     * @hidden
     * Hide from docs
     *  ------
     *
     * Gets list of downloads for current user
     * @param callback Callback that will be triggered post downloads load
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getFileDownloads(callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw new Error('[files.getFileDownloads] Callback cannot be null');
        }
        sendMessageToParent('files.getFileDownloads', [], callback);
    }
    files_1.getFileDownloads = getFileDownloads;
    /**
     * @hidden
     * Hide from docs
     *
     * Open download preference folder if fileObjectId value is undefined else open folder containing the file with id fileObjectId
     * @param fileObjectId - Id of the file whose containing folder should be opened
     * @param callback Callback that will be triggered post open download folder/path
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function openDownloadFolder(fileObjectId, callback) {
        if (fileObjectId === void 0) { fileObjectId = undefined; }
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw new Error('[files.openDownloadFolder] Callback cannot be null');
        }
        sendMessageToParent('files.openDownloadFolder', [fileObjectId], callback);
    }
    files_1.openDownloadFolder = openDownloadFolder;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates add 3P cloud storage provider flow, where a pop up window opens for user to select required
     * 3P provider from the configured policy supported 3P provider list, following which user authentication
     * for selected 3P provider is performed on success of which the selected 3P provider support is added for user
     * @beta
     *
     * @param callback Callback that will be triggered post add 3P cloud storage provider action.
     * If the error is encountered (and hence passed back), no provider value is sent back.
     * For success scenarios, error value will be passed as null and a valid provider value is sent.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function addCloudStorageProvider(callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.addCloudStorageProvider] callback cannot be null');
        }
        sendMessageToParent('files.addCloudStorageProvider', [], callback);
    }
    files_1.addCloudStorageProvider = addCloudStorageProvider;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates signout of 3P cloud storage provider flow, which will remove the selected
     * 3P cloud storage provider from the list of added providers. No other user input and / or action
     * is required except the 3P cloud storage provider to signout from
     *
     * @param logoutRequest 3P cloud storage provider remove action request content
     * @param callback Callback that will be triggered post signout of 3P cloud storage provider action
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function removeCloudStorageProvider(logoutRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.removeCloudStorageProvider] callback cannot be null');
        }
        if (!(logoutRequest && logoutRequest.content)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.removeCloudStorageProvider] 3P cloud storage provider request content is missing');
        }
        sendMessageToParent('files.removeCloudStorageProvider', [logoutRequest], callback);
    }
    files_1.removeCloudStorageProvider = removeCloudStorageProvider;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates the add 3P cloud storage file flow, which will add a new file for the given 3P provider
     *
     * @param addNewFileRequest 3P cloud storage provider add action request content
     * @param callback Callback that will be triggered post adding a new file flow is finished
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function addCloudStorageProviderFile(addNewFileRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.addCloudStorageProviderFile] callback cannot be null');
        }
        if (!(addNewFileRequest && addNewFileRequest.content)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.addCloudStorageProviderFile] 3P cloud storage provider request content is missing');
        }
        sendMessageToParent('files.addCloudStorageProviderFile', [addNewFileRequest], callback);
    }
    files_1.addCloudStorageProviderFile = addCloudStorageProviderFile;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates the rename 3P cloud storage file flow, which will rename an existing file in the given 3P provider
     *
     * @param renameFileRequest 3P cloud storage provider rename action request content
     * @param callback Callback that will be triggered post renaming an existing file flow is finished
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function renameCloudStorageProviderFile(renameFileRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.renameCloudStorageProviderFile] callback cannot be null');
        }
        if (!(renameFileRequest && renameFileRequest.content)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.renameCloudStorageProviderFile] 3P cloud storage provider request content is missing');
        }
        sendMessageToParent('files.renameCloudStorageProviderFile', [renameFileRequest], callback);
    }
    files_1.renameCloudStorageProviderFile = renameCloudStorageProviderFile;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates the delete 3P cloud storage file(s) / folder (folder has to be empty) flow,
     * which will delete existing file(s) / folder from the given 3P provider
     *
     * @param deleteFileRequest 3P cloud storage provider delete action request content
     * @param callback Callback that will be triggered post deleting existing file(s) flow is finished
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function deleteCloudStorageProviderFile(deleteFileRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.deleteCloudStorageProviderFile] callback cannot be null');
        }
        if (!(deleteFileRequest &&
            deleteFileRequest.content &&
            deleteFileRequest.content.itemList &&
            deleteFileRequest.content.itemList.length > 0)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.deleteCloudStorageProviderFile] 3P cloud storage provider request content details are missing');
        }
        sendMessageToParent('files.deleteCloudStorageProviderFile', [deleteFileRequest], callback);
    }
    files_1.deleteCloudStorageProviderFile = deleteCloudStorageProviderFile;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates the download 3P cloud storage file(s) flow,
     * which will download existing file(s) from the given 3P provider in the teams client side without sharing any file info in the callback
     *
     * @param downloadFileRequest 3P cloud storage provider download file(s) action request content
     * @param callback Callback that will be triggered post downloading existing file(s) flow is finished
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function downloadCloudStorageProviderFile(downloadFileRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.downloadCloudStorageProviderFile] callback cannot be null');
        }
        if (!(downloadFileRequest &&
            downloadFileRequest.content &&
            downloadFileRequest.content.itemList &&
            downloadFileRequest.content.itemList.length > 0)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.downloadCloudStorageProviderFile] 3P cloud storage provider request content details are missing');
        }
        sendMessageToParent('files.downloadCloudStorageProviderFile', [downloadFileRequest], callback);
    }
    files_1.downloadCloudStorageProviderFile = downloadCloudStorageProviderFile;
    /**
     * @hidden
     * Hide from docs
     *
     * Initiates the upload 3P cloud storage file(s) flow, which will upload file(s) to the given 3P provider
     * @beta
     *
     * @param uploadFileRequest 3P cloud storage provider upload file(s) action request content
     * @param callback Callback that will be triggered post uploading file(s) flow is finished
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function uploadCloudStorageProviderFile(uploadFileRequest, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!callback) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.uploadCloudStorageProviderFile] callback cannot be null');
        }
        if (!(uploadFileRequest &&
            uploadFileRequest.content &&
            uploadFileRequest.content.itemList &&
            uploadFileRequest.content.itemList.length > 0)) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.uploadCloudStorageProviderFile] 3P cloud storage provider request content details are missing');
        }
        if (!uploadFileRequest.content.destinationFolder) {
            throw getSdkError(ErrorCode.INVALID_ARGUMENTS, '[files.uploadCloudStorageProviderFile] Invalid destination folder details');
        }
        sendMessageToParent('files.uploadCloudStorageProviderFile', [uploadFileRequest], callback);
    }
    files_1.uploadCloudStorageProviderFile = uploadCloudStorageProviderFile;
    /**
     * @hidden
     * Hide from docs
     *
     * Register a handler to be called when a user's 3P cloud storage provider list changes i.e.
     * post adding / removing a 3P provider, list is updated
     *
     * @param handler - When 3P cloud storage provider list is updated this handler is called
     *
     * @internal Limited to Microsoft-internal use
     */
    function registerCloudStorageProviderListChangeHandler(handler) {
        ensureInitialized(runtime);
        if (!handler) {
            throw new Error('[registerCloudStorageProviderListChangeHandler] Handler cannot be null');
        }
        registerHandler('files.cloudStorageProviderListChange', handler);
    }
    files_1.registerCloudStorageProviderListChangeHandler = registerCloudStorageProviderListChangeHandler;
    /**
     * @hidden
     * Hide from docs
     *
     * Register a handler to be called when a user's 3P cloud storage provider content changes i.e.
     * when file(s) is/are added / renamed / deleted / uploaded, the list of files is updated
     *
     * @param handler - When 3P cloud storage provider content is updated this handler is called
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerCloudStorageProviderContentChangeHandler(handler) {
        ensureInitialized(runtime);
        if (!handler) {
            throw new Error('[registerCloudStorageProviderContentChangeHandler] Handler cannot be null');
        }
        registerHandler('files.cloudStorageProviderContentChange', handler);
    }
    files_1.registerCloudStorageProviderContentChangeHandler = registerCloudStorageProviderContentChangeHandler;
    function getSdkError(errorCode, message) {
        var sdkError = {
            errorCode: errorCode,
            message: message,
        };
        return sdkError;
    }
})(files || (files = {}));

;// CONCATENATED MODULE: ./src/private/meetingRoom.ts





/**
 * @hidden
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var meetingRoom;
(function (meetingRoom) {
    /**
     * @hidden
     * Fetch the meeting room info that paired with current client.
     *
     * @returns Promise resolved with meeting room info or rejected with SdkError value
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getPairedMeetingRoomInfo() {
        return new Promise(function (resolve) {
            ensureInitialized(runtime);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            resolve(sendAndHandleSdkError('meetingRoom.getPairedMeetingRoomInfo'));
        });
    }
    meetingRoom.getPairedMeetingRoomInfo = getPairedMeetingRoomInfo;
    /**
     * @hidden
     * Send a command to paired meeting room.
     *
     * @param commandName The command name.
     * @returns Promise resolved upon completion or rejected with SdkError value
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function sendCommandToPairedMeetingRoom(commandName) {
        return new Promise(function (resolve) {
            if (!commandName || commandName.length == 0) {
                throw new Error('[meetingRoom.sendCommandToPairedMeetingRoom] Command name cannot be null or empty');
            }
            ensureInitialized(runtime);
            if (!isSupported()) {
                throw errorNotSupportedOnPlatform;
            }
            resolve(sendAndHandleSdkError('meetingRoom.sendCommandToPairedMeetingRoom', commandName));
        });
    }
    meetingRoom.sendCommandToPairedMeetingRoom = sendCommandToPairedMeetingRoom;
    /**
     * @hidden
     * Registers a handler for meeting room capabilities update.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler The handler to invoke when the capabilities of meeting room update.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerMeetingRoomCapabilitiesUpdateHandler(handler) {
        if (!handler) {
            throw new Error('[meetingRoom.registerMeetingRoomCapabilitiesUpdateHandler] Handler cannot be null');
        }
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('meetingRoom.meetingRoomCapabilitiesUpdate', function (capabilities) {
            ensureInitialized(runtime);
            handler(capabilities);
        });
    }
    meetingRoom.registerMeetingRoomCapabilitiesUpdateHandler = registerMeetingRoomCapabilitiesUpdateHandler;
    /**
     * @hidden
     * Hide from docs
     * Registers a handler for meeting room states update.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler The handler to invoke when the states of meeting room update.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerMeetingRoomStatesUpdateHandler(handler) {
        if (!handler) {
            throw new Error('[meetingRoom.registerMeetingRoomStatesUpdateHandler] Handler cannot be null');
        }
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('meetingRoom.meetingRoomStatesUpdate', function (states) {
            ensureInitialized(runtime);
            handler(states);
        });
    }
    meetingRoom.registerMeetingRoomStatesUpdateHandler = registerMeetingRoomStatesUpdateHandler;
    /**
     * @hidden
     *
     * Checks if the meetingRoom capability is supported by the host
     * @returns boolean to represent whether the meetingRoom capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.meetingRoom ? true : false;
    }
    meetingRoom.isSupported = isSupported;
})(meetingRoom || (meetingRoom = {}));

;// CONCATENATED MODULE: ./src/private/notifications.ts




var notifications;
(function (notifications) {
    /**
     * @hidden
     * display notification API.
     *
     * @param message - Notification message.
     * @param notificationType - Notification type
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function showNotification(showNotificationParameters) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('notifications.showNotification', [showNotificationParameters]);
    }
    notifications.showNotification = showNotification;
    /**
     * @hidden
     *
     * Checks if the notifications capability is supported by the host
     * @returns boolean to represent whether the notifications capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.notifications ? true : false;
    }
    notifications.isSupported = isSupported;
})(notifications || (notifications = {}));

;// CONCATENATED MODULE: ./src/private/remoteCamera.ts





/**
 * @hidden
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var remoteCamera;
(function (remoteCamera) {
    /**
     * @hidden
     * Enum used to indicate possible camera control commands.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var ControlCommand;
    (function (ControlCommand) {
        ControlCommand["Reset"] = "Reset";
        ControlCommand["ZoomIn"] = "ZoomIn";
        ControlCommand["ZoomOut"] = "ZoomOut";
        ControlCommand["PanLeft"] = "PanLeft";
        ControlCommand["PanRight"] = "PanRight";
        ControlCommand["TiltUp"] = "TiltUp";
        ControlCommand["TiltDown"] = "TiltDown";
    })(ControlCommand = remoteCamera.ControlCommand || (remoteCamera.ControlCommand = {}));
    /**
     * @hidden
     * Enum used to indicate the reason for the error.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var ErrorReason;
    (function (ErrorReason) {
        ErrorReason[ErrorReason["CommandResetError"] = 0] = "CommandResetError";
        ErrorReason[ErrorReason["CommandZoomInError"] = 1] = "CommandZoomInError";
        ErrorReason[ErrorReason["CommandZoomOutError"] = 2] = "CommandZoomOutError";
        ErrorReason[ErrorReason["CommandPanLeftError"] = 3] = "CommandPanLeftError";
        ErrorReason[ErrorReason["CommandPanRightError"] = 4] = "CommandPanRightError";
        ErrorReason[ErrorReason["CommandTiltUpError"] = 5] = "CommandTiltUpError";
        ErrorReason[ErrorReason["CommandTiltDownError"] = 6] = "CommandTiltDownError";
        ErrorReason[ErrorReason["SendDataError"] = 7] = "SendDataError";
    })(ErrorReason = remoteCamera.ErrorReason || (remoteCamera.ErrorReason = {}));
    /**
     * @hidden
     * Enum used to indicate the reason the session was terminated.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var SessionTerminatedReason;
    (function (SessionTerminatedReason) {
        SessionTerminatedReason[SessionTerminatedReason["None"] = 0] = "None";
        SessionTerminatedReason[SessionTerminatedReason["ControlDenied"] = 1] = "ControlDenied";
        SessionTerminatedReason[SessionTerminatedReason["ControlNoResponse"] = 2] = "ControlNoResponse";
        SessionTerminatedReason[SessionTerminatedReason["ControlBusy"] = 3] = "ControlBusy";
        SessionTerminatedReason[SessionTerminatedReason["AckTimeout"] = 4] = "AckTimeout";
        SessionTerminatedReason[SessionTerminatedReason["ControlTerminated"] = 5] = "ControlTerminated";
        SessionTerminatedReason[SessionTerminatedReason["ControllerTerminated"] = 6] = "ControllerTerminated";
        SessionTerminatedReason[SessionTerminatedReason["DataChannelError"] = 7] = "DataChannelError";
        SessionTerminatedReason[SessionTerminatedReason["ControllerCancelled"] = 8] = "ControllerCancelled";
        SessionTerminatedReason[SessionTerminatedReason["ControlDisabled"] = 9] = "ControlDisabled";
        SessionTerminatedReason[SessionTerminatedReason["ControlTerminatedToAllowOtherController"] = 10] = "ControlTerminatedToAllowOtherController";
    })(SessionTerminatedReason = remoteCamera.SessionTerminatedReason || (remoteCamera.SessionTerminatedReason = {}));
    /**
     * @hidden
     * Fetch a list of the participants with controllable-cameras in a meeting.
     *
     * @param callback - Callback contains 2 parameters, error and participants.
     * error can either contain an error of type SdkError, incase of an error, or null when fetch is successful
     * participants can either contain an array of Participant objects, incase of a successful fetch or null when it fails
     * participants: object that contains an array of participants with controllable-cameras
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getCapableParticipants(callback) {
        if (!callback) {
            throw new Error('[remoteCamera.getCapableParticipants] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('remoteCamera.getCapableParticipants', callback);
    }
    remoteCamera.getCapableParticipants = getCapableParticipants;
    /**
     * @hidden
     * Request control of a participant's camera.
     *
     * @param participant - Participant specifies the participant to send the request for camera control.
     * @param callback - Callback contains 2 parameters, error and requestResponse.
     * error can either contain an error of type SdkError, incase of an error, or null when fetch is successful
     * requestResponse can either contain the true/false value, incase of a successful request or null when it fails
     * requestResponse: True means request was accepted and false means request was denied
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function requestControl(participant, callback) {
        if (!participant) {
            throw new Error('[remoteCamera.requestControl] Participant cannot be null');
        }
        if (!callback) {
            throw new Error('[remoteCamera.requestControl] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('remoteCamera.requestControl', [participant], callback);
    }
    remoteCamera.requestControl = requestControl;
    /**
     * @hidden
     * Send control command to the participant's camera.
     *
     * @param ControlCommand - ControlCommand specifies the command for controling the camera.
     * @param callback - Callback to invoke when the command response returns.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function sendControlCommand(ControlCommand, callback) {
        if (!ControlCommand) {
            throw new Error('[remoteCamera.sendControlCommand] ControlCommand cannot be null');
        }
        if (!callback) {
            throw new Error('[remoteCamera.sendControlCommand] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('remoteCamera.sendControlCommand', [ControlCommand], callback);
    }
    remoteCamera.sendControlCommand = sendControlCommand;
    /**
     * @hidden
     * Terminate the remote  session
     *
     * @param callback - Callback to invoke when the command response returns.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function terminateSession(callback) {
        if (!callback) {
            throw new Error('[remoteCamera.terminateSession] Callback cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('remoteCamera.terminateSession', callback);
    }
    remoteCamera.terminateSession = terminateSession;
    /**
     * @hidden
     * Registers a handler for change in participants with controllable-cameras.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the list of participants with controllable-cameras changes.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerOnCapableParticipantsChangeHandler(handler) {
        if (!handler) {
            throw new Error('[remoteCamera.registerOnCapableParticipantsChangeHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('remoteCamera.capableParticipantsChange', handler);
    }
    remoteCamera.registerOnCapableParticipantsChangeHandler = registerOnCapableParticipantsChangeHandler;
    /**
     * @hidden
     * Registers a handler for error.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when there is an error from the camera handler.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerOnErrorHandler(handler) {
        if (!handler) {
            throw new Error('[remoteCamera.registerOnErrorHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('remoteCamera.handlerError', handler);
    }
    remoteCamera.registerOnErrorHandler = registerOnErrorHandler;
    /**
     * @hidden
     * Registers a handler for device state change.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the controlled device changes state.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerOnDeviceStateChangeHandler(handler) {
        if (!handler) {
            throw new Error('[remoteCamera.registerOnDeviceStateChangeHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('remoteCamera.deviceStateChange', handler);
    }
    remoteCamera.registerOnDeviceStateChangeHandler = registerOnDeviceStateChangeHandler;
    /**
     * @hidden
     * Registers a handler for session status change.
     * Only one handler can be registered at a time. A subsequent registration replaces an existing registration.
     *
     * @param handler - The handler to invoke when the current session status changes.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerOnSessionStatusChangeHandler(handler) {
        if (!handler) {
            throw new Error('[remoteCamera.registerOnSessionStatusChangeHandler] Handler cannot be null');
        }
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('remoteCamera.sessionStatusChange', handler);
    }
    remoteCamera.registerOnSessionStatusChangeHandler = registerOnSessionStatusChangeHandler;
    /**
     * @hidden
     *
     * Checks if the remoteCamera capability is supported by the host
     * @returns boolean to represent whether the remoteCamera capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.remoteCamera ? true : false;
    }
    remoteCamera.isSupported = isSupported;
})(remoteCamera || (remoteCamera = {}));

;// CONCATENATED MODULE: ./src/private/appEntity.ts





/**
 * @hidden
 * Namespace to interact with the application entities specific part of the SDK.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var appEntity;
(function (appEntity_1) {
    /**
     * @hidden
     * Hide from docs
     * --------
     * Open the Tab Gallery and retrieve the app entity
     * @param threadId ID of the thread where the app entity will be created
     * @param categories A list of application categories that will be displayed in the opened tab gallery
     * @param subEntityId An object that will be made available to the application being configured
     *                      through the Context's subEntityId field.
     * @param callback Callback that will be triggered once the app entity information is available.
     *                 The callback takes two arguments: an SdkError in case something happened (i.e.
     *                 no permissions to execute the API) and the app entity configuration, if available
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function selectAppEntity(threadId, categories, subEntityId, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!threadId || threadId.length == 0) {
            throw new Error('[appEntity.selectAppEntity] threadId name cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[appEntity.selectAppEntity] Callback cannot be null');
        }
        sendMessageToParent('appEntity.selectAppEntity', [threadId, categories, subEntityId], callback);
    }
    appEntity_1.selectAppEntity = selectAppEntity;
    /**
     * @hidden
     *
     * Checks if the appEntity capability is supported by the host
     * @returns boolean to represent whether the appEntity capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.appEntity ? true : false;
    }
    appEntity_1.isSupported = isSupported;
})(appEntity || (appEntity = {}));

;// CONCATENATED MODULE: ./src/private/teams.ts







/**
 * @hidden
 * Namespace to interact with the `teams` specific part of the SDK.
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var teams;
(function (teams) {
    var ChannelType;
    (function (ChannelType) {
        ChannelType[ChannelType["Regular"] = 0] = "Regular";
        ChannelType[ChannelType["Private"] = 1] = "Private";
        ChannelType[ChannelType["Shared"] = 2] = "Shared";
    })(ChannelType = teams.ChannelType || (teams.ChannelType = {}));
    /**
     * @hidden
     * Get a list of channels belong to a Team
     *
     * @param groupId - a team's objectId
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function getTeamChannels(groupId, callback) {
        ensureInitialized(runtime, FrameContexts.content);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!groupId) {
            throw new Error('[teams.getTeamChannels] groupId cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[teams.getTeamChannels] Callback cannot be null');
        }
        sendMessageToParent('teams.getTeamChannels', [groupId], callback);
    }
    teams.getTeamChannels = getTeamChannels;
    /**
     * @hidden
     * Allow 1st party apps to call this function when they receive migrated errors to inform the Hub/Host to refresh the siteurl
     * when site admin renames siteurl.
     *
     * @param threadId - ID of the thread where the app entity will be created; if threadId is not
     * provided, the threadId from route params will be used.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function refreshSiteUrl(threadId, callback) {
        ensureInitialized(runtime);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!threadId) {
            throw new Error('[teams.refreshSiteUrl] threadId cannot be null or empty');
        }
        if (!callback) {
            throw new Error('[teams.refreshSiteUrl] Callback cannot be null');
        }
        sendMessageToParent('teams.refreshSiteUrl', [threadId], callback);
    }
    teams.refreshSiteUrl = refreshSiteUrl;
    /**
     * @hidden
     *
     * Checks if teams capability is supported by the host
     * @returns boolean to represent whether the teams capability is supported
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        return ensureInitialized(runtime) && runtime.supports.teams ? true : false;
    }
    teams.isSupported = isSupported;
    /**
     * @hidden
     * @internal
     * Limited to Microsoft-internal use
     */
    var fullTrust;
    (function (fullTrust) {
        /**
         * @hidden
         * @internal
         * Limited to Microsoft-internal use
         */
        var joinedTeams;
        (function (joinedTeams) {
            /**
             * @hidden
             * Allows an app to retrieve information of all user joined teams
             *
             * @param teamInstanceParameters - Optional flags that specify whether to scope call to favorite teams
             * @returns Promise that resolves with information about the user joined teams or rejects with an error when the operation has completed
             *
             * @internal
             * Limited to Microsoft-internal use
             */
            function getUserJoinedTeams(teamInstanceParameters) {
                return new Promise(function (resolve) {
                    ensureInitialized(runtime);
                    if (!isSupported()) {
                        throw errorNotSupportedOnPlatform;
                    }
                    if ((GlobalVars.hostClientType === HostClientType.android ||
                        GlobalVars.hostClientType === HostClientType.teamsRoomsAndroid ||
                        GlobalVars.hostClientType === HostClientType.teamsPhones ||
                        GlobalVars.hostClientType === HostClientType.teamsDisplays) &&
                        !isCurrentSDKVersionAtLeast(getUserJoinedTeamsSupportedAndroidClientVersion)) {
                        var oldPlatformError = { errorCode: ErrorCode.OLD_PLATFORM };
                        throw new Error(JSON.stringify(oldPlatformError));
                    }
                    /* eslint-disable-next-line strict-null-checks/all */ /* Fix tracked by 5730662 */
                    resolve(sendAndUnwrap('getUserJoinedTeams', teamInstanceParameters));
                });
            }
            joinedTeams.getUserJoinedTeams = getUserJoinedTeams;
            /**
             * @hidden
             *
             * Checks if teams.fullTrust.joinedTeams capability is supported by the host
             * @returns boolean to represent whether the teams.fullTrust.joinedTeams capability is supported
             *
             * @throws Error if {@linkcode app.initialize} has not successfully completed
             *
             * @internal
             * Limited to Microsoft-internal use
             */
            function isSupported() {
                return ensureInitialized(runtime) && runtime.supports.teams
                    ? runtime.supports.teams.fullTrust
                        ? runtime.supports.teams.fullTrust.joinedTeams
                            ? true
                            : false
                        : false
                    : false;
            }
            joinedTeams.isSupported = isSupported;
        })(joinedTeams = fullTrust.joinedTeams || (fullTrust.joinedTeams = {}));
        /**
         * @hidden
         * Allows an app to get the configuration setting value
         *
         * @param key - The key for the config setting
         * @returns Promise that resolves with the value for the provided configuration setting or rejects with an error when the operation has completed
         *
         * @internal
         * Limited to Microsoft-internal use
         */
        function getConfigSetting(key) {
            return new Promise(function (resolve) {
                ensureInitialized(runtime);
                if (!isSupported()) {
                    throw errorNotSupportedOnPlatform;
                }
                resolve(sendAndUnwrap('getConfigSetting', key));
            });
        }
        fullTrust.getConfigSetting = getConfigSetting;
        /**
         * @hidden
         *
         * Checks if teams.fullTrust capability is supported by the host
         * @returns boolean to represent whether the teams.fullTrust capability is supported
         *
         * @throws Error if {@linkcode app.initialize} has not successfully completed
         *
         * @internal
         * Limited to Microsoft-internal use
         */
        function isSupported() {
            return ensureInitialized(runtime) && runtime.supports.teams
                ? runtime.supports.teams.fullTrust
                    ? true
                    : false
                : false;
        }
        fullTrust.isSupported = isSupported;
    })(fullTrust = teams.fullTrust || (teams.fullTrust = {}));
})(teams || (teams = {}));

;// CONCATENATED MODULE: ./src/private/videoEx.ts
var videoEx_awaiter = (undefined && undefined.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var videoEx_generator = (undefined && undefined.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};









/**
 * @hidden
 * Extended video API
 * @beta
 *
 * @internal
 * Limited to Microsoft-internal use
 */
var videoEx;
(function (videoEx) {
    var videoPerformanceMonitor = inServerSideRenderingEnvironment()
        ? undefined
        : new VideoPerformanceMonitor(sendMessageToParent);
    /**
     * @hidden
     * Error level when notifying errors to the host, the host will decide what to do acording to the error level.
     * @beta
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    var ErrorLevel;
    (function (ErrorLevel) {
        ErrorLevel["Fatal"] = "fatal";
        ErrorLevel["Warn"] = "warn";
    })(ErrorLevel = videoEx.ErrorLevel || (videoEx.ErrorLevel = {}));
    /**
     * @hidden
     * Register to process video frames
     * @beta
     *
     * @param parameters - Callbacks and configuration to process the video frames. A host may support either {@link VideoFrameHandler} or {@link VideoBufferHandler}, but not both.
     * To ensure the video effect works on all supported hosts, the video app must provide both {@link VideoFrameHandler} and {@link VideoBufferHandler}.
     * The host will choose the appropriate callback based on the host's capability.
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerForVideoFrame(parameters) {
        var _this = this;
        var _a, _b;
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        if (!parameters.videoFrameHandler || !parameters.videoBufferHandler) {
            throw new Error('Both videoFrameHandler and videoBufferHandler must be provided');
        }
        if (ensureInitialized(runtime, FrameContexts.sidePanel)) {
            registerHandler('video.setFrameProcessTimeLimit', function (timeLimit) { return videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.setFrameProcessTimeLimit(timeLimit); }, false);
            if ((_a = runtime.supports.video) === null || _a === void 0 ? void 0 : _a.mediaStream) {
                registerHandler('video.startVideoExtensibilityVideoStream', function (mediaStreamInfo) { return videoEx_awaiter(_this, void 0, void 0, function () {
                    var streamId, metadataInTexture, generator, _a;
                    var _b, _c;
                    return videoEx_generator(this, function (_d) {
                        switch (_d.label) {
                            case 0:
                                streamId = mediaStreamInfo.streamId, metadataInTexture = mediaStreamInfo.metadataInTexture;
                                if (!metadataInTexture) return [3 /*break*/, 2];
                                return [4 /*yield*/, processMediaStreamWithMetadata(streamId, parameters.videoFrameHandler, notifyError, videoPerformanceMonitor)];
                            case 1:
                                _a = _d.sent();
                                return [3 /*break*/, 4];
                            case 2: return [4 /*yield*/, processMediaStream(streamId, parameters.videoFrameHandler, notifyError, videoPerformanceMonitor)];
                            case 3:
                                _a = _d.sent();
                                _d.label = 4;
                            case 4:
                                generator = _a;
                                // register the video track with processed frames back to the stream
                                !inServerSideRenderingEnvironment() &&
                                    ((_c = (_b = window['chrome']) === null || _b === void 0 ? void 0 : _b.webview) === null || _c === void 0 ? void 0 : _c.registerTextureStream(streamId, generator));
                                return [2 /*return*/];
                        }
                    });
                }); }, false);
                sendMessageToParent('video.mediaStream.registerForVideoFrame', [parameters.config]);
            }
            else if ((_b = runtime.supports.video) === null || _b === void 0 ? void 0 : _b.sharedFrame) {
                registerHandler('video.newVideoFrame', function (videoBufferData) {
                    if (videoBufferData) {
                        videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportStartFrameProcessing(videoBufferData.width, videoBufferData.height);
                        var timestamp_1 = videoBufferData.timestamp;
                        parameters.videoBufferHandler(normalizedVideoBufferData(videoBufferData), function () {
                            videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.reportFrameProcessed();
                            notifyVideoFrameProcessed(timestamp_1);
                        }, notifyError);
                    }
                }, false);
                sendMessageToParent('video.registerForVideoFrame', [parameters.config]);
            }
            else {
                // should not happen if isSupported() is true
                throw errorNotSupportedOnPlatform;
            }
            videoPerformanceMonitor === null || videoPerformanceMonitor === void 0 ? void 0 : videoPerformanceMonitor.startMonitorSlowFrameProcessing();
        }
    }
    videoEx.registerForVideoFrame = registerForVideoFrame;
    function normalizedVideoBufferData(videoBufferData) {
        videoBufferData['videoFrameBuffer'] = videoBufferData['videoFrameBuffer'] || videoBufferData['data'];
        delete videoBufferData['data'];
        return videoBufferData;
    }
    /**
     * @hidden
     * Video extension should call this to notify host that the current selected effect parameter changed.
     * If it's pre-meeting, host will call videoEffectCallback immediately then use the videoEffect.
     * If it's the in-meeting scenario, we will call videoEffectCallback when apply button clicked.
     * @beta
     * @param effectChangeType - the effect change type.
     * @param effectId - Newly selected effect id. {@linkcode VideoEffectCallBack}
     * @param effectParam Variant for the newly selected effect. {@linkcode VideoEffectCallBack}
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function notifySelectedVideoEffectChanged(effectChangeType, effectId, effectParam) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('video.videoEffectChanged', [effectChangeType, effectId, effectParam]);
    }
    videoEx.notifySelectedVideoEffectChanged = notifySelectedVideoEffectChanged;
    /**
     * @hidden
     * Register the video effect callback, host uses this to notify the video extension the new video effect will by applied
     * @beta
     * @param callback - The VideoEffectCallback to invoke when registerForVideoEffect has completed
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function registerForVideoEffect(callback) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        registerHandler('video.effectParameterChange', createEffectParameterChangeCallback(callback, videoPerformanceMonitor), false);
        sendMessageToParent('video.registerForVideoEffect');
    }
    videoEx.registerForVideoEffect = registerForVideoEffect;
    /**
     * @hidden
     * Send personalized effects to Teams client
     * @beta
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function updatePersonalizedEffects(effects) {
        ensureInitialized(runtime, FrameContexts.sidePanel);
        if (!video.isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        sendMessageToParent('video.personalizedEffectsChanged', [effects]);
    }
    videoEx.updatePersonalizedEffects = updatePersonalizedEffects;
    /**
     * @hidden
     *
     * Checks if video capability is supported by the host
     * @beta
     *
     * @throws Error if {@linkcode app.initialize} has not successfully completed
     *
     * @returns boolean to represent whether the video capability is supported
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function isSupported() {
        ensureInitialized(runtime);
        return video.isSupported();
    }
    videoEx.isSupported = isSupported;
    /**
     * @hidden
     * Sending notification to host finished the video frame processing, now host can render this video frame
     * or pass the video frame to next one in video pipeline
     * @beta
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function notifyVideoFrameProcessed(timestamp) {
        sendMessageToParent('video.videoFrameProcessed', [timestamp]);
    }
    /**
     * @hidden
     * Sending error notification to host
     * @beta
     * @param errorMessage - The error message that will be sent to the host
     * @param errorLevel - The error level that will be sent to the host
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function notifyError(errorMessage, errorLevel) {
        if (errorLevel === void 0) { errorLevel = ErrorLevel.Warn; }
        sendMessageToParent('video.notifyError', [errorMessage, errorLevel]);
    }
    /**
     * @hidden
     * Sending fatal error notification to host. Call this function only when your app meets fatal error and can't continue.
     * The host will stop the video pipeline and terminate this session, and optionally, show an error message to the user.
     * @beta
     * @param errorMessage - The error message that will be sent to the host
     *
     * @internal
     * Limited to Microsoft-internal use
     */
    function notifyFatalError(errorMessage) {
        ensureInitialized(runtime);
        if (!video.isSupported()) {
            throw errorNotSupportedOnPlatform;
        }
        notifyError(errorMessage, ErrorLevel.Fatal);
    }
    videoEx.notifyFatalError = notifyFatalError;
})(videoEx || (videoEx = {}));

;// CONCATENATED MODULE: ./src/private/index.ts












;// CONCATENATED MODULE: ./src/index.ts



})();

/******/ 	return __webpack_exports__;
/******/ })()
;
});
//# sourceMappingURL=MicrosoftTeams.js.map