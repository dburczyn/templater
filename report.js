var sizeOf = require('image-size');
var express = require('express')
var bodyParser = require('body-parser')
var app = express()
const {
  encode,
  decode
} = require('base64-arraybuffer');
app.use(bodyParser.json({
  limit: '50mb'
}));
app.post('/', async (req, res) => {
  data = req.body.template;
  sampledata = JSON.parse(req.body.data);
  sampledata.images = req.body.images;
  width = 1;
  height = 1;

  // sampledata.images.forEach(element => {
  //    fs.writeFile("/tmp/test" + sampledata.images.length, sampledata.images, function(err) {

  //     if(err) {
  //         return console.log(err);
  //     }

  //     console.log("The file was saved!");
  // });
  // });



  const rap = await createReportWithImg(sampledata)

  res.send(rap);
})
app.listen(3000)


var createReport = require('docx-templates');
const qrcode = require('yaqrcode');

function prepareData(sampledata) {
  var preparedData = {};
  preparedData.name = sampledata['ado:publishing'].model._name;
  preparedData.class = sampledata['ado:publishing'].model._class;
  preparedData.images = sampledata.images;
  preparedData.chapters = [];
  for (let [index, val] of sampledata['ado:publishing'].model.notebook.chapter.entries()) {
    var chapter = {}
    chapter.name = val._name
    chapter.attributes = []
    var attributes = [].concat(val.attribute || []).concat(val.relation || []).concat(typeof val.group !== 'undefined' ? val.group.relation || [] : []).concat(typeof val.group !== 'undefined' ? val.group.attribute || [] : [])
    for (let [aindex, aval] of attributes.entries()) {
      var attri = {};
      if (typeof aval._name !== 'undefined') {
        attri.name = aval._name;
        attri.value = getValue(aval)
        attri.searchname = aval._idname;
        chapter.attributes.push(attri);
      } else if (typeof aval._class !== 'undefined') {
        attri.name = aval._class;
        attri.searchname = aval._idclass;
        attri.value = getValue(aval)
        chapter.attributes.push(attri);
      } else {
        attri.name = "noname";
        attri.value = "noval";
        chapter.attributes.push(attri);
      }
    }
    preparedData.chapters[index] = chapter
  }
  preparedData.objects = [];
  for (let [oindex, oval] of sampledata['ado:publishing'].model.object.entries()) {
    var object = {}
    object.name = oval._name;
    object.class = oval._class;
    object.ochapters = [];
    /////////////////////////////////////////////////////////////////////////
    for (let [index, val] of oval.notebook.chapter.entries()) {
      var ochapter = {}
      ochapter.name = val._name
      ochapter.oattributes = []
      var oattributes = [].concat(val.attribute || []).concat(val.relation || []).concat(typeof val.group !== 'undefined' ? val.group.relation || [] : []).concat(typeof val.group !== 'undefined' ? val.group.attribute || [] : [])

      for (let [aindex, aval] of oattributes.entries()) {
        var oattri = {};
        if (typeof aval._name !== 'undefined') {
          oattri.name = aval._name;
          oattri.value = getValue(aval)
          oattri.searchname = aval._idname;
          ochapter.oattributes.push(oattri);
        } else if (typeof aval._class !== 'undefined') {
          oattri.name = aval._class;
          oattri.searchname = aval._idclass;
          oattri.value = getValue(aval)
          ochapter.oattributes.push(oattri);
        } else {
          oattri.name = "noname";
          oattri.value = "noval";
          ochapter.attributes.push(oattri);
        }
      }
      object.ochapters[index] = ochapter
    }
    //////////////////////////////////////////////////////////////
    preparedData.objects[oindex] = object;
  }
  return preparedData;
};

function getComplexVals(passedval, passedinp) {
  var vals = []
  if (typeof passedinp.complexvalues.member !== 'undefined' && typeof passedinp.complexvalues.member[0] === 'undefined') {
    for (let [index2, val2] of passedinp.complexvalues.member.complexvalues.member.entries()) {
      if (val2._name === passedval._name) {
        vals.push(getValue(val2))
      }
    }
  }
  if (typeof passedinp.complexvalues.member[0] !== 'undefined') {
    for (let [index, val] of passedinp.complexvalues.member.entries()) {
      for (let [index2, val2] of val.complexvalues.member.entries()) {
        if (val2._name === passedval._name) {
          vals.push(getValue(val2))
        }
      }
    }
  }
  return vals;
}

function getComplex(inp) {
  var complexarray = []
  for (let [index, val] of inp.columns.member.entries()) {
    var complexmember = {}
    complexmember.name = val._name;
    complexmember.values = getComplexVals(val, inp)
    complexarray.push(complexmember)
  }
  return complexarray;
}

function getValue(inp) {
  if (typeof inp.attrval !== 'undefined') {
    if (inp.attrval.attrvaltype._type === 'ENUM') {
      return (inp.attrval._name);
    } else if (inp.attrval.attrvaltype._type === 'BOOL') {
      return (inp.attrval.value === '0' ? "no" : "yes");
    } else if (inp.attrval.attrvaltype._type === 'LONGSTRING' || inp.attrval.attrvaltype._type === 'UNSIGNED INTEGER' || inp.attrval.attrvaltype._type === 'FILE_POINTER' || inp.attrval.attrvaltype._type === 'SHORTSTRING') {
      if (typeof inp.attrval.value.p !== 'undefined') {
        return inp.attrval.value.p
      } else return (inp.attrval.value);
    } else if (inp.attrval.attrvaltype._type === 'ADOSTRING' || inp.attrval.attrvaltype._type === 'STRING') {
      return (typeof inp.attrval.value.p === 'undefined' ? inp.attrval.value : inp.attrval.value.p);
    } else if (inp.attrval.attrvaltype._type === 'DOUBLE' || inp.attrval.attrvaltype._type === 'UTC') {
      return (inp.attrval["alternate-value"]);
    } else if (inp.attrval.attrvaltype._type === 'INTERREF') {
      return (inp.attrval.relation.link.endpoint._name);
    } else {
      return inp.attrval.attrvaltype._type;
    }
  } else if (typeof inp.link !== 'undefined' && typeof inp.link.endpoint !== 'undefined') {
    return (inp.link.endpoint._name);
  } else if (typeof inp.link !== 'undefined' && typeof inp.link[1] !== 'undefined') {
    var ret = '';
    for (let [index, val] of inp.link.entries()) {
      ret = ret + val.endpoint._name + '\n';
    }
    return ret.replace(/^\s+|\s+$/g, "");
  } else if (inp._complex === '1') {
    return getComplex(inp);
  } else {
    return "wrongtype";
  }
}

function gN(obj, searched) {
  for (const prop in obj) {
    const value = obj[prop];
    if (typeof value === 'object') {
      var sub = gN(value, searched);
      if (sub) {
        return sub;
      }
    } else {
      if (value === searched) {
        return obj._name;
      }
    }
  }
  return null;
}

function gV(obj, searched) {
  for (const prop in obj) {
    const value = obj[prop];
    if (typeof value === 'object') {
      var sub = gV(value, searched);
      if (sub) {
        return sub;
      }
    } else {
      if (value === searched) {
        //  return getValue(obj);
        return obj.value;
      }
    }
  }
  return null;
}

async function createReportWithImg(sampledata) {
  const prepareddata = {}
  prepareddata.model = await prepareData(sampledata);

  function toBuffer(ab) {
    var buf = Buffer.alloc(ab.byteLength);
    var view = new Uint8Array(ab);
    for (var i = 0; i < buf.length; ++i) {
      buf[i] = view[i];
    }
    return buf;
  }
  const template = await toBuffer(decode(data));

  const report = await createReport({
    template,
    data: prepareddata,
    cmdDelimiter: ['{', '}'],
    noSandbox: true,
    additionalJsContext: {
      qr: async (contents) => {
        const dataUrl = await qrcode(contents, {
          size: 500
        });
        const data = await dataUrl.slice('data:image/gif;base64,'.length);
        return {
          width: 6,
          height: 6,
          data,
          extension: '.gif'
        };
      },
      insertImg: function (image,w) {
        var img = Buffer.from(image, 'base64');
        var dimensions = sizeOf(img);
        var ratiow = dimensions.width / width < 1 ? dimensions.width / width : 1;
        var ratioh = dimensions.height/dimensions.width;
        // var ratioh = dimensions.height / height < 1 ? dimensions.height / height : 1;
        width = dimensions.width;
        // height = dimensions.height;
        if (ratiow !== 1) {
          width = 1
        }
        // if (ratioh !== 1) {
        //   height = 1
        // }
        return {
          width: w * ratiow,
          height: w*ratiow*ratioh,
          data: image,
          extension: '.png'
        };
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
      srt: (inp, by) => {
        return inp.sort(function (a, b) {
          return ('' + gV(a, by)).localeCompare(gV(b, by));
        })
      },
      fE: (inp) => {
        function isEmpty(value) {
          return value.attributes.length > 0;
        }
        return inp.filter(isEmpty);
      },
      foE: (inp) => {
        function isEmpty(value) {
          return value.oattributes.length > 0;
        }
        return inp.filter(isEmpty);
      },
      fT: (inp, searched) => {
        function filT(value) {


          for (let [index, val] of searched.entries()) {
            if (value.name === val)
              return false;
          }
          return true;



        }
        return inp.filter(filT);
      },
      gCN: (inp) => {
        if (typeof inp !== 'undefined') {
          return inp.name
        } else return "";
      },
      gCV: (inp) => {
        if (typeof inp !== 'undefined') {
          return "vvv"
        } else return "";
      },
      mI: (inp) => {
        var maxlength = 0
        var outarray = []
        for (let [index, val] of inp.entries()) {
          if (maxlength < val.values.length) {
            maxlength = val.values.length
          }
        }
        for (var i = 0; i < maxlength; i++) {
          var inarray = []
          outarray.push(inarray)
        }
        for (let [index, val] of inp.entries()) {
          for (let [index2, val2] of val.values.entries()) {
            outarray[index2].push(val2)
          }
        }
        return outarray;
      },
      gN: gN,
      gV: gV
    }
  });


  const rs = await encode(report);
  return rs;

}