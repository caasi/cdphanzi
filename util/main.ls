console.log opentype

font-path = '../font/'
eudc = 'cdpeudc.tte'
text = for i from 5797 to 5813
  String.from-char-code 57344 + i

err, font <-! opentype.load "#{font-path}#{eudc}"
if err
  console.log "Fail to load #{eudc}: #{err}"
else
  ctx = document.get-element-by-id(\canvas)get-context \2d
  path = font.get-path text.join(''), 0, 50, 32
  path.draw ctx
