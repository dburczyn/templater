var createReport = require('docx-templates');
var sizeOf = require('image-size');
var express = require('express');
var bodyParser = require('body-parser');
const {
  encode,
  decode
} = require('base64-arraybuffer');
const port = 3000;
var app = express();
app.use(bodyParser.json({
  limit: '100mb'
}));
app.post('/', async (req, res) => {
  data = req.body.template;
  sampledata = JSON.parse(req.body.data);
  sampledata.images = req.body.images;
  width = 1;
  try {
    const rap = await createReportWithImg(sampledata);
    res.send(rap);
  } catch (e) {
    console.log(e);
    res.status(418).send(e);
  }
});
app.listen(port, () => console.log(`App listening on port ${port}!`));
function getAttrs(val){
  const result = [];
  function getAttrsInn(obj) {
      for (const prop in obj) {
           const value = obj[prop];
           if (typeof value === 'object') {
           getAttrsInn(value);
           }
           else if (prop==='_controltype')
           {
              result.push(obj);
           }
      }
  }
  getAttrsInn(val);
  return result;
}
function prepareData(sampledata) {
  var preparedData = {};
  if (sampledata['ado:publishing'].hasOwnProperty('model')) {
    preparedData.name = sampledata['ado:publishing'].model._name;
    preparedData.class = sampledata['ado:publishing'].model._class;
    preparedData.images = sampledata.images;
    preparedData.chapters = [];
    for (let [index, val] of sampledata['ado:publishing'].model.notebook.chapter.entries()) {
      var chapter = {};
      chapter.name = val._name;
      chapter.attributes = [];
      var attributes = getAttrs(val);
      for (let [i,aval] of attributes.entries()) {
        var attri = {};
        if (aval.hasOwnProperty('_name')) {
          attri.name = aval._name;
          attri.value = getValue(aval);
          attri.searchname = aval._idname;
          if(attri.searchname!=='INSTANCE_CHANGE_HISTORY'){
          chapter.attributes.push(attri);
          }
        } else if (aval.hasOwnProperty('_class')) {
          attri.name = aval._class;
          attri.searchname = aval._idclass;
          attri.value = getValue(aval);
            chapter.attributes.push(attri);
        } else {
          attri.name = "noname";
          attri.value = "noval";
            chapter.attributes.push(attri);
        }
      }
      preparedData.chapters[index] = chapter;
    }
    preparedData.objects = [];
    for (let [oindex, oval] of sampledata['ado:publishing'].model.object.entries()) {
      var object = {};
      object.name = oval._name;
      object.class = oval._class;
      object.ochapters = [];
      for (let [index, val] of oval.notebook.chapter.entries()) {
        var ochapter = {};
        ochapter.name = val._name;
        ochapter.attributes = [];
        var attributes = getAttrs(val);
        for (let [i,aval] of attributes.entries()) {
          var oattri = {};
          if (aval.hasOwnProperty('_name')) {
            oattri.name = aval._name;
            oattri.value = getValue(aval);
            oattri.searchname = aval._idname;
            if(oattri.searchname!=='INSTANCE_CHANGE_HISTORY'){
            ochapter.attributes.push(oattri);
            }
          } else if (aval.hasOwnProperty('_class')) {
            oattri.name = aval._class;
            oattri.searchname = aval._idclass;
            oattri.value = getValue(aval);
            ochapter.attributes.push(oattri);
          } else {
            oattri.name = "noname";
            oattri.value = "noval";
            ochapter.attributes.push(oattri);
          }
        }
        object.ochapters[index] = ochapter;
      }
            preparedData.objects[oindex] = object;
    }
    preparedData.objects.sort(function (a, b) {
      return (''+a.name).localeCompare((''+b.name));
    });
  } else if (sampledata['ado:publishing'].hasOwnProperty('object')) {
    preparedData.name = sampledata['ado:publishing'].object._name;
    preparedData.class = sampledata['ado:publishing'].object._class;
    preparedData.images = "";
    preparedData.chapters = [];
    for (let [index, val] of sampledata['ado:publishing'].object.notebook.chapter.entries()) {
      var sochapter = {};
      sochapter.name = val._name;
      sochapter.attributes = [];
      var attributes = getAttrs(val);
       for (let [i,aval] of attributes.entries()) {
        var soattri = {};
        if (aval.hasOwnProperty('_name')) {
          soattri.name = aval._name;
          soattri.value = getValue(aval);
          soattri.searchname = aval._idname;
          sochapter.attributes.push(soattri);
        } else if (aval.hasOwnProperty('_class')) {
          soattri.name = aval._class;
          soattri.searchname = aval._idclass;
          soattri.value = getValue(aval);
          sochapter.attributes.push(soattri);
        } else {
         soattri.name = "noname";
         soattri.value = "noval";
          sochapter.attributes.push(soattri);
        }
      }
      preparedData.chapters[index] = sochapter;
    }
    preparedData.objects = [];
  }
  return preparedData;
}
function getComplexVals(passedval, passedinp) {
  var vals = [];
  if(passedinp.hasOwnProperty('complexvalues') && passedinp.complexvalues.hasOwnProperty('member')){
  if (!Array.isArray(passedinp.complexvalues.member)) {
    for (let [i,val2] of passedinp.complexvalues.member.complexvalues.member.entries()) {
      if (val2._name === passedval._name) {
        vals.push(getValue(val2));
      }
    }
  }
  else if (Array.isArray(passedinp.complexvalues.member)) {
    for (let [i,val] of passedinp.complexvalues.member.entries()) {
      for (let [i,val2] of val.complexvalues.member.entries()) {
        if (val2._name === passedval._name) {
          vals.push(getValue(val2));
        }
      }
    }
  }
}
  return vals;
}
function getComplex(inp) {
if (inp.hasOwnProperty('columns')){
  var complexarray = [];
  for (let [i,val] of inp.columns.member.entries()) {
    var complexmember = {};
    complexmember.name = val._name;
    complexmember.values = getComplexVals(val, inp);
    complexarray.push(complexmember);
  }
  return complexarray;
}
}
function getValue(inp) {
  if (inp.hasOwnProperty('attrval')) {
    if (inp.attrval.attrvaltype._type === 'ENUM') {
      return (inp.attrval._name);
    } else if (inp.attrval.attrvaltype._type === 'BOOL') {
      return (inp.attrval.value === 0 ? "nie" : "tak");
    } else if (inp.attrval.attrvaltype._type === 'LONGSTRING' || inp.attrval.attrvaltype._type === 'UNSIGNED INTEGER' || inp.attrval.attrvaltype._type === 'FILE_POINTER' || inp.attrval.attrvaltype._type === 'SHORTSTRING'|| inp.attrval.attrvaltype._type === 'INTEGER' || inp.attrval.attrvaltype._type === 'ADOSTRING' || inp.attrval.attrvaltype._type === 'STRING') {
      return (inp.attrval.value.hasOwnProperty('p') ? inp.attrval.value.p : inp.attrval.value );
      }
      else if (inp.attrval.attrvaltype._type === 'DOUBLE' || inp.attrval.attrvaltype._type === 'UTC') {
      return (inp.attrval["alternate-value"]||inp.attrval.value);
    } else if (inp.attrval.attrvaltype._type === 'INTERREF') {
      if (inp.attrval.relation.hasOwnProperty('link'))
      {
        return (inp.attrval.relation.link.endpoint._name);
      }
    } else {
      return inp.attrval.attrvaltype._type;
    }
  }
  else if(inp._type=== 'FILE_POINTER')
  {
    var splited = inp.value.param.split('/');
    return splited[splited.length-1];
  }

  else if (inp.hasOwnProperty('link') && inp.link.hasOwnProperty('endpoint')) {
    return (inp.link.endpoint._name);
  } else if (inp.hasOwnProperty('link') && Array.isArray(inp.link)) {
    var ret = '';
    for (let [i,val] of inp.link.entries()) {
      ret = ret +'\u2022'+ val.endpoint._name + '\n';
    }
    return ret.replace(/^\s+|\s+$/g, "");
  } else if (inp._complex === 1) {
    return getComplex(inp);
  } else {
    return "";
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
  const prepareddata = {};
  prepareddata.model = await prepareData(sampledata);
   function toArray(obj) {
    for (const prop in obj) {
        const value = obj[prop];
        if (typeof value === 'object') {
            toArray(value);
        }
        else if (typeof value === 'string') {
          obj[prop]=value.replace(/&amp;/g, '&')
          .replace(/&quot;/g, '"')
          .replace(/&gt;/g, '>')
          .replace(/&lt;/g, '<')
          .replace(/&#039;/g, "'");
        }
    }
}
toArray(prepareddata);
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
    additionalJsContext: {
      insertImg: function (image, w) {
        var img = Buffer.from(image, 'base64');
        var dimensions = sizeOf(img);
        var ratiow = dimensions.width / width < 1 ? dimensions.width / width : 1;
        var ratioh = dimensions.height / dimensions.width;
        width = dimensions.width;
        if (ratiow !== 1) {
          width = 1;
        }
        return {
          width: w * ratiow,
          height: w * ratiow * ratioh,
          data: image,
          extension: '.png'
        };
      },
      srt: (inp, by) => {
        if (typeof inp !== 'undefined') {
          return inp.sort(function (a, b) {
            return ('' + gV(a, by)).localeCompare(gV(b, by));
          });
        }
        return;
      },
      fE: (inp) => {
        function isEmpty(value) {
          return value.attributes.length > 0;
        }
        return inp.filter(isEmpty);
      },
      fT: (inp, searched) => {
        function filT(value) {
          for (let [i,val] of searched.entries()) {
            if (value.name === val) {
              return false;
            }
          }
          return true;
        }
        return inp.filter(filT);
      },
      gCN: (inp) => {
        if (typeof inp !== 'undefined') {
          return inp.name;
        } else {
          return "";
        }
      },
      mI: (inp) => {
        var maxlength = 0;
        var outarray = [];
        for (let [i,val] of inp.entries()) {
          if (maxlength < val.values.length) {
            maxlength = val.values.length;
          }
        }
        for (var i = 0; i < maxlength; i++) {
          var inarray = [];
          outarray.push(inarray);
        }
        for (let [i,val] of inp.entries()) {
          for (let [index2, val2] of val.values.entries()) {
            outarray[index2].push(val2);
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
