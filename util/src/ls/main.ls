dim = 22
font-path = '../font/'
fonts =
  * 'cdpeudc.tte'
    'cdpeudck.tte'
    'nulleudc.tte'
eudc = fonts.0
text =
  for i from 0xA000 to 0xF000 by 0x1000
    for j from 0x000 to 0xFF0 by 0x010
      str = for k from 0x0 to 0xF
        String.from-char-code i + j + k
      str.join ''

err, font <-! opentype.load "#{font-path}#{eudc}"
if err
  console.log "Fail to load #{eudc}: #{err}"
else
  console.log font
  canvas = $(\#canvas).0
  canvas
    ..width  = text.length * dim * 16
    ..height = text.0.length * dim + dim * 0.8
  ctx = canvas.get-context \2d
  for i, page of text
    for j, row of page
      path = font.get-path row, +i * dim * 16, +j * dim + dim * 0.8, dim
      path.draw ctx
