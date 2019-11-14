var createReport = require('docx-templates');
var sampledata = require('./sampleadonis');
var statics = require('./statics');

sampledata.statics = statics;

createReport({
  template: 'raportStandardowy.docx',
  output: 'test.docx',
  data: sampledata,
  additionalJsContext: {
    injectSvg: (inp) => {
      return { width: 16, height: 6, path: './'+inp.substring(1, inp.length-1)+'.png' };
    },
    getVal: (inp) => {
      // return  Object.keys(inp) ;
      if (typeof inp.attrval!=='undefined' )
      {return  (inp.attrval.value) ;}
     else {return inp.attrval;}
    }
  }
});