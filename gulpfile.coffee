'use strict'


taskMasterOptions = 
  dirname: 'src/gulp' 
  pattern: '*.coffee' 
  cwd: process.cwd() 
  watchExt: '.watch'  

gulp = require('gulp-task-master')(taskMasterOptions)


gulp.task "watch",  ['compile-main.watch']
gulp.task "bot", ['compile-main.watch']



gulp.task "default", ['compile-main']

