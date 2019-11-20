var createReport = require('docx-templates');
var sampledata = require('./sampleadonis');
var statics = require('./statics');
sampledata.statics = statics;
const qrcode = require('yaqrcode');
function getVal(inp){
  if (typeof inp.attrval !== 'undefined') {
    if (inp.attrval.attrvaltype._type === 'ENUM') {
      return (inp.attrval._name);
    } else if (inp.attrval.attrvaltype._type === 'BOOL') {
      return (inp.attrval.value === '0' ? "nie" : "tak");
    } else if (inp.attrval.attrvaltype._type === 'LONGSTRING' || inp.attrval.attrvaltype._type === 'UNSIGNED INTEGER' || inp.attrval.attrvaltype._type === 'FILE_POINTER') {
      return (inp.attrval.value);
    } else if (inp.attrval.attrvaltype._type === 'ADOSTRING' || inp.attrval.attrvaltype._type === 'STRING') {
      return (typeof inp.attrval.value.p === 'undefined' ? inp.attrval.value : inp.attrval.value.p);
    } else if (inp.attrval.attrvaltype._type === 'DOUBLE' || inp.attrval.attrvaltype._type === 'UTC') {
      return (inp.attrval["alternate-value"]);
    }  else if (inp.attrval.attrvaltype._type === 'INTERREF') {
      return (inp.attrval.relation.link.endpoint._name);
    }
    else {
      return inp.attrval.attrvaltype._type;
    }
  } else if (typeof inp.link !== 'undefined' && typeof inp.link.endpoint !== 'undefined') {
    return (inp.link.endpoint._name);
  } else if (inp._complex ==='1') {
     return getComplex(inp);
  }
  else {
    return "wrongtype";
  }
}
function gN(obj,searched) {
	for (const prop in obj) {
		const value = obj[prop];
		if (typeof value === 'object') {
			var sub =gN(value,searched);
			if (sub)
			{return sub;}
		}
		else {
			if (value===searched)
			{
			 return obj._name;
			}
		}
	}
  return null;
}
function gV(obj,searched) {
	for (const prop in obj) {
		const value = obj[prop];
		if (typeof value === 'object') {
			var sub =gV(value,searched);
			if (sub)
			{return sub;}
		}
		else {
			if (value===searched)
			{
			 return getVal(obj);
			}
		}
	}
  return null;
}
async function createReportWithImg() {
await createReport({
  template: 'raportNew.docx',
  output: 'test.docx',
  data: sampledata,
  // cmdDelimiter: '---',
noSandbox: true,
  additionalJsContext: {
    qr: async (contents) => {
      const dataUrl = await qrcode(contents, { size: 500 });
      const data = await dataUrl.slice('data:image/gif;base64,'.length);
      return { width: 6, height: 6, data, extension: '.gif' };
    },
    insertImg: function(url) {
      data=url.slice('data:image/gif;base64,'.length);
      return { width: 15, height: 6, data, extension: '.png' };
      // return new Promise(function(resolve) {
      //      var image = new Canvas.Image;
      //      image.onload = function() {
      //     //   var canvas = new Canvas(image.width, image.height);
      //     //     var ctx = canvas.getContext("2d");
      //     //     const imageDPI = 96;  // Put all your images with the same DPI
      //     //     canvas.height = this.naturalHeight;
      //     //     canvas.width = this.naturalWidth;
      //     //     ctx.drawImage(this, 0, 0);
      //     //     var dataUrl = canvas.toDataURL("image/png", 1);
      //         resolve({
      //             // height: (this.naturalHeight * 2.54) / imageDPI,
      //             // width: (this.naturalWidth * 2.54) / imageDPI,
      //             height:5,
      //             width:6,
      //             data: image.src.slice(
      //                 "data:image/png;base64,".length
      //             ),
      //             extension: ".png",
      //         });
      //      };
      //     image.src = url;
      // });
  },
    checkNonempty: (inp) => {
      if (typeof inp.attribute !== 'undefined' || typeof inp.relation !== 'undefined') {
        return true;
      } else {
        return false;
      }
    },
    getChapterName: (inp) => {
      {
        return (inp._name);
      }
    },
    checkArray: (inp) => {
      if ((typeof inp.attribute !== 'undefined' && typeof inp.attribute[0] !== 'undefined') || (typeof inp.relation !== 'undefined' && typeof inp.relation[0] !== 'undefined')) {
        return true;
      } else {
        return false;
      }
    },
    checkSingle: (inp) => {
      if ((typeof inp.attribute !== 'undefined' && typeof inp.attribute[0] === 'undefined') || (typeof inp.relation !== 'undefined' && typeof inp.relation[0] === 'undefined')) {
        return true;
      } else {
        return false;
      }
    },
    getType: (inp) => {
      if (typeof inp.attribute !== 'undefined' && typeof inp.attribute[0] !== 'undefined') {
        return inp.attribute;
      } else if (typeof inp.relation !== 'undefined' && typeof inp.relation[0] !== 'undefined') {
        return inp.relation;
      }
    },
    getName: (inp) => {
      if (typeof inp._name !== 'undefined') {
        return inp._name;
      } else if (typeof inp._class !== 'undefined') {
        return inp._class;
      } else {
        return "noname";
      }
    },
    getSingleName: (inp) => {
      if (typeof inp.attribute !== 'undefined' && typeof inp.attribute._name !== 'undefined') {
        return inp.attribute._name;
      } else if (typeof inp.relation !== 'undefined' && typeof inp.relation._class !== 'undefined') {
        return inp.relation._class;
      } else {
        return "noname";
      }
    },
    getSingleVal: (inp) => {
      if (typeof inp.attribute !== 'undefined' && typeof inp.attribute.attrval !== 'undefined') {
        return (inp.attribute.attrval.attrvaltype._type);
      } else if (typeof inp.relation !== 'undefined' && typeof inp.relation.link !== 'undefined' && typeof inp.relation.link.endpoint !== 'undefined') {
        return (inp.relation.link.endpoint._name);
      } else {
        return "wrongtype";
      }
    },
    getComplexName: (inp) => {
  return ["ahnhntnhnghhnnhnhnnhnhhn","gnhnhnhnhnhnhgnhnhb","chnhnhnhnhghnh"]
    },
    getComplexValue: (inp) => {
      return [["ahnhntnhnghhnnhnhnnhnhhn","gnhnhnhnhnhnhgnhnhb","chnhnhnhnhghnh"],["yhere6h","wefwveg3","uo987kk9657"],["4hhehe","dfgfdgdf","ee344y"]]
        },
        gN: gN,
        gV:gV
      }
});
}
createReportWithImg();