/**
 *  Google Spreadsheets API for CommonJS
 */

var xml2js = require("xml2js-expat"),
    request = require("request"),
    _ = require("underscore");

/**
 *  A Google Spreadsheet.
 *
 *  @param {Object} opts      Object containing at least the spreadsheet key. Can also contain authFactory for producing authenticated requests
 *  @constructor
 */
function Spreadsheet(opts) {
  //Backwards compatibility
  if (!(opts instanceof Object)) {
    opts = {
      key: opts
    };
  }
  if (!opts.key) {
    throw new Error("A Spreadsheet must have a key.");
  }
  this.opts = opts;
}

/**
 *  Retrieve all worksheets of the Spreadsheet.
 *
 *  @param {Function} fn     Function called for each Worksheet: `function(err,worksheet){}`
 *  @return Itself for chaining.
 */
Spreadsheet.prototype.worksheets = function(fn) {
  var spreadsheet = this;
  return this.load(function(err, atom) {
    if (err) return fn(err);
    spreadsheet.author = atom.author;
    spreadsheet.title = atom.title["#"];
    spreadsheet.updated = new Date(atom.updated);
    spreadsheet.sheetCount = parseInt(atom["openSearch:totalResults"], 10);
    spreadsheet.startIndex = parseInt(atom["openSearch:startIndex"], 10);
    // Create a worksheet from each entry
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    entries.forEach(function(entry, i) {
      fn(null, new Worksheet(spreadsheet.opts, spreadsheet, entry.id.match(/\w+$/)[0], i + 1, entry.title["#"], new Date(entry.updated)));
    });
  });
};

/**
 *  Retrieve all worksheet array of the Spreadsheet.
 *
 *  @param {Function} fn     Function called for each Worksheet: `function(err,spreadsheet, worksheets){}`
 *  @return Itself for chaining.
 */
Spreadsheet.prototype.worksheetArray = function(fn) {
  var spreadsheet = this;
  return this.load(function(err, atom) {
    if (err) return fn(err);
    spreadsheet.author = atom.author;
    spreadsheet.title = atom.title["#"];
    spreadsheet.updated = new Date(atom.updated);
    spreadsheet.startIndex = parseInt(atom["openSearch:startIndex"], 10);
    spreadsheet.sheetCount = parseInt(atom["openSearch:totalResults"], 10);
    // Create a worksheet array from each entry
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    var sheets = [],
      len = entries.length;
    for (var i = 0; i < len; i += 1) {
      var entry = entries[i];
      sheets.push(new Worksheet(spreadsheet.opts, spreadsheet, entry.id.match(/\w+$/)[0], i + 1, entry.title["#"], new Date(entry.updated)));
    }
    fn(null, spreadsheet, sheets);
  });
};

/**
 *  Retrieve a worksheet by worksheetID.
 *
 *  @param {String} id       The ID of the worksheet.
 *  @param {Function} fn     Function called with a worksheet instance.
 *  @return Itself for chaining.
 */
Spreadsheet.prototype.worksheet = function(id, fn) {
  var spreadsheet = this;
  return this.load(function(err, atom) {
    if (err) return fn(err);
    spreadsheet.author = atom.author;
    spreadsheet.title = atom.title["#"];
    spreadsheet.updated = new Date(atom.updated);
    spreadsheet.startIndex = parseInt(atom["openSearch:startIndex"], 10);
    spreadsheet.sheetCount = parseInt(atom["openSearch:totalResults"], 10);
    // Create a worksheet array from each entry
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    if (toString.call(id) == "[object Number]") {
      if (id > 0 && id <= entries.length) {
        var entry = entries[id - 1];
        return fn(null, new Worksheet(spreadsheet.opts, spreadsheet, entry.id.match(/\w+$/)[0], id, entry.title["#"], new Date(entry.updated)));
      }
    } else {
      var len = entries.length;
      for (var i = 0; i < len; i += 1) {
        var entry = entries[i];
        var eId = entry.id.match(/\w+$/)[0];
        if (eId == id) {
          return fn(null, new Worksheet(spreadsheet.opts, spreadsheet, eId, i + 1, entry.title["#"], new Date(entry.updated)));
        }
      }
    }
    fn(new Error("A Worksheet must have an valid id."));
  });
};

/**
 *  Loads the Spreadsheet from Google Data API.
 *
 *  @param {Function} fn     Called when loaded.
 *  @return Itself for chaining.
 *  @private
 */
Spreadsheet.prototype.load = function(fn) {
  atom(["https://spreadsheets.google.com/feeds/worksheets", this.opts.key, this.opts.authFactory?"private":"public", "values"].join("/"), this.opts, fn);
  return this;
};

/**
 *  Each Spreadsheet contains at least one worksheet.
 *
 *  @param {String|Number} id  The worksheetID or page number of the worksheet.
 *  @constructor
 */
function Worksheet(opts, spreadsheet, id, index, title, updated) {
  this.opts = opts;
  if (!spreadsheet) {
    throw new Error("A Worksheet must belong to a Spreadsheet.");
  }
  if (!id) {
    throw new Error("A Worksheet must have an id.");
  }
  this.id = id;
  this.spreadsheet = spreadsheet;
  this.index = typeof id == "number" ? id : index;
  this.title = title;
  this.updated = updated;
}


/**
 *  Go through each row in the worksheet.
 *
 *  @param {Function} fn     Function called for each row: function(err,row,meta){}
 *  @return Itself for chaining.
 */
Worksheet.prototype.eachRow = function(fn) {
  return this.load("list", function(err, atom) {
    if (err) return fn(err);
    if (!atom.entry) return fn(new Error("No rows found."));
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    entries.forEach(function(entry, i) {
      fn(null, Row(entry), Meta(entry, atom, i));
    })
  })
};

/**
 *  Finds a single Row in a Worksheet by id.
 *
 *  @param {String} id       The ID of the single row (returned in the meta of an eachRow callback)
 *  @param {Function} fn     Function called for each row: function(err,row,meta){}
 *  @return Itself for chaining.
 */
Worksheet.prototype.row = function(id, fn) {
  return this.load("list", id, function(err, entry) {
    if (err) return fn(err);
    if (!entry) return fn(new Error("No row found with id:" + id));
    fn(null, Row(entry), Meta(entry));
  });
};

/**
 *  Go through each cell in the worksheet.
 *
 *  @param {Function} fn     Function called for each cell: function(err,cell,meta){}
 *  @return Itself for chaining.
 */
Worksheet.prototype.eachCell = function(fn) {
  var ws = this;
  return this.load("cells", function(err, atom) {
    if (err) return fn(err);
    var entries = Array.isArray(atom.entry) ? atom.entry : [atom.entry];
    console.log(entries);
    entries.forEach(function(entry, i) {
      fn(null, Cell(entry), Meta(entry, atom, i));
    });
  });
};

/**
 * Looks through an iterative for cell ids, replacing them with the cell values.
 * A modifier can optionally be supplied to control the replacement value; it will be passed the Cell object.
 * Cell IDs should be in the following format: R1C5 where R1 is row 1 and C5 is column 5, for example.
 * 
 * @param  {Object/Array} map   Object/Array to iterate over.
 * @param  {Function} fn        Callback to be passed the new object
 * @param  {Function} modifier  Optional modifier function
 * @return Itself for chaining
 */
Worksheet.prototype.mapCells = function(map, fn, modifier) {

  var mapCopy = _.clone(map);

  //Filter out valid unique cell IDs
  var cellIds = [];
  walk(map, function(value) {
    if (_.isString(value) && /^R\d+?C\d+?$/.test(value)) {
      cellIds.push(value);
    }
  });

  var cells = {};

  //Retrieve each cell
  _.each(cellIds, function(cellId) {

    this.cell(cellId, function(err, cell, meta) {
      cells[cellId] = cell;

      //Once all have been collected, remap values
      if (_.keys(cells).length == cellIds.length) {

        walk(mapCopy, function(value) {
          if (_.has(cells, value)) {
            if (modifier) {
              return modifier(cells[value]);
            } else {
              return cells[value];
            }
          }
        });

        fn(mapCopy);
      }

    });
  }, this);

  return this;

};

/**
 *  Finds a single Cell in a Worksheet by id.
 *
 *  @param {String} id       The ID of the single cell (returned in the meta of an eachRell callback)
 *  @param {Function} fn     Function called for each cell: function(err,cell,meta){}
 *  @return Itself for chaining.
 */
Worksheet.prototype.cell = function(id, fn) {
  return this.load("cells", id, function(err, entry) {
    if (err) return fn(err);
    if (!entry) return fn(new Error("No cell found with id:" + id));
    fn(null, Cell(entry), Meta(entry));
  });
};

/**
 *  Loads the Worksheet from Google Data API.
 *
 *  @param {String} type     Type of feed. Should be "list" or "cells".
 *  @param {Function} fn     Called when loaded.
 *  @return Itself for chaining.
 *  @private
 */
Worksheet.prototype.load = function(type, id, fn) {
  if (typeof id == "function") {
    fn = id;
    id = "";
  }
  else id = "/" + id;
  atom(["https://spreadsheets.google.com/feeds", type, this.opts.key, this.id, this.opts.authFactory?"private":"public", "values"].join("/") + id, this.opts, fn);
  return this;
};

/**
 *  Instance of a Row
 *
 *  @param {Object} entry    An Atom Entry to extract the rows fields from
 */
function Row(entry) {
  if (!(this instanceof Row)) {
    return new Row(entry);
  }
  // Build a row object from each entrys gdx: items.
  for (var k in entry) {
    if (k.indexOf("gsx:") === 0) {
      this[k.slice(4)] = cellValue(entry[k]);
    }
  }
}


/**
 *  Instance of a Cell
 *
 *  @param {Object} entry    An Atom Entry to extract the cell contents from
 */
function Cell(entry) {
  if (!(this instanceof Cell)) {
    return new Cell(entry);
  }
  this.row = parseInt(entry["gs:cell"]["@"].row, 10);
  this.col = parseInt(entry["gs:cell"]["@"].col, 10);
  this.value = cellValue(entry["gs:cell"]["#"]);
}

/**
 *  Instance of a Meta
 *  @private
 */
function Meta(entry, atom, i) {
  if (!(this instanceof Meta)) {
    return new Meta(entry, atom, i);
  }
  // And a meta object from the entrys "id","updated", "index" and "total"
  this.id = entry.id;
  this.updated = new Date(entry.updated);
  if (atom) {
    this.index = parseInt(atom["openSearch:startIndex"], 10) + i;
    this.total = parseInt(atom["openSearch:totalResults"], 10);
  }
}

/**
 *  Helper method for retrieving entries from an atom feed.
 *
 *  @param {String} url      The URL to the atom feed.
 *  @param {Function} fn     A callback like: function(err,atom){}
 *  @private
 */
function atom(url, opts, fn) {

  var options = {
    url: url + "?hl=en",
    method: "GET",
    headers: {}
  };

  var run = function() {
    request(options, function(err, res, body) {

      //If request error, return
      if (err) return fn(err);

      //If 401 and authfactory is present, get a new header indicating that we want to refresh the token and then run again
      if (res.statusCode == 401 && opts.authFactory) {
        opts.authFactory.getAuthHeader(true, function(header) {
          options.authorization = header;
          run();
        });
      }

      //Else if 200, request was successull
      else if (res.statusCode == 200) {
        var parser = new xml2js.Parser();
        parser.on("end", function(root) {
          fn(null, root);
        }).on("error", function(err) {
          fn(err);
        }).parse(body);
      }

      //Otherwise throw error
      else {
        fn(new Error(body));
      }
    });
  };

  //If an auth factory has been passed in the main options, use it for the authorization header
  if (opts.authFactory) {
    opts.authFactory.getAuthHeader(function(header) {
      options.headers.Authorization = header;
      run();
    });
  } else {
    run();
  }

}

/**
 *  Helper method for converting cell value to null or not if it is empty.
 *
 *  @param {Object} obj      The value of cell.
 *  @private
 */
function cellValue(obj) {
  if (obj === null) return null;
  if (toString.call(obj) == "[object Array]" || toString.call(obj) == "[object String]") return (obj.length === 0) ? null : obj;
  for (var key in obj) {
    if (hasOwnProperty.call(obj, key)) return obj;
  }
  return null;
}

/**
 * Recursively walk over an object/array, applying an interator function.
 * If the iterator returns a defined value, that value will be applied to the object/array.
 * 
 * @param  {Object/Array} object  Object to walk over. Can be any mixture of iteratives (e.g. mix of arrays and objects)
 * @param  {Function} iterator    Function to call on each non-iterative value. If it returns a value, the value will be applied to the object
 * @private
 */
function walk(object, iterator) {
  for (var key in object) {
    var value = object[key];
    if (value instanceof Object || value instanceof Array) {
      walk(value, iterator);
    } else {
      var newValue = iterator(value, key);
      if (newValue !== undefined) {
        object[key] = newValue;
      }
    }
  }
}

module.exports = Spreadsheet;

// Expose the other constructors for testing
Spreadsheet.Worksheet = Worksheet;
Spreadsheet.Row = Row;
Spreadsheet.Cell = Cell;
Spreadsheet.Meta = Meta;