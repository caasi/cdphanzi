(function(){
  var dim, fontPath, fonts, eudc, text, res$, i$, i, lresult$, j$, j, str, res1$, k$, k;
  dim = 22;
  fontPath = '../font/';
  fonts = ['cdpeudc.tte', 'cdpeudck.tte', 'nulleudc.tte'];
  eudc = fonts[0];
  res$ = [];
  for (i$ = 0xA000; i$ <= 0xF000; i$ += 0x1000) {
    i = i$;
    lresult$ = [];
    for (j$ = 0x000; j$ <= 0xFF0; j$ += 0x010) {
      j = j$;
      res1$ = [];
      for (k$ = 0x0; k$ <= 0xF; ++k$) {
        k = k$;
        res1$.push(String.fromCharCode(i + j + k));
      }
      str = res1$;
      lresult$.push(str.join(''));
    }
    res$.push(lresult$);
  }
  text = res$;
  opentype.load(fontPath + "" + eudc, function(err, font){
    var canvas, x$, ctx, i, ref$, page, j, row, path;
    if (err) {
      console.log("Fail to load " + eudc + ": " + err);
    } else {
      console.log(font);
      canvas = $('#canvas')[0];
      x$ = canvas;
      x$.width = text.length * dim * 16;
      x$.height = text[0].length * dim + dim * 0.8;
      ctx = canvas.getContext('2d');
      for (i in ref$ = text) {
        page = ref$[i];
        for (j in page) {
          row = page[j];
          path = font.getPath(row, +i * dim * 16, +j * dim + dim * 0.8, dim);
          path.draw(ctx);
        }
      }
    }
  });
}).call(this);
