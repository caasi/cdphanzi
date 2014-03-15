require! <[gulp gulp-concat gulp-livereload node-static tiny-lr]>
gutil             = require \gulp-util
livescript        = require \gulp-livescript
jade              = require \gulp-jade
livereload-server = tiny-lr!
livereload        = -> gulp-livereload livereload-server

path =
  root:  '../'
  build: './'

gulp.task \js ->
  gulp.src do
    * 'src/ls/*.ls'
    ...
  .pipe gulp-concat 'main.ls'
  .pipe livescript!
  .pipe gulp.dest path.build
  .pipe livereload!

gulp.task \html ->
  gulp.src 'src/index.jade'
  .pipe jade!
  .pipe gulp.dest path.build
  .pipe livereload!

gulp.task \build <[js html]>

gulp.task \static (next) ->
  server = new node-static.Server path.root
  port = 8888
  require \http .createServer (req, res) !->
    req.addListener(\end -> server.serve req, res)resume!
  .listen port, !->
    gutil.log "Server listening on port: #{gutil.colors.magenta port}"
    next!

gulp.task \watch ->
  gulp.watch 'src/ls/*.ls'    <[js]>
  gulp.watch 'src/index.jade' <[html]>

gulp.task \livereload ->
  port = 35729
  livereload-server.listen port, ->
    return gulp.log it if it
    gutil.log "LiveReload listening on port: #{gutil.colors.magenta port}"

gulp.task \default <[build static watch livereload]>
