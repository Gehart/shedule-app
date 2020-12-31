'use strict';

const { dest } = require('gulp');

var gulp         = require('gulp'),                 // Подключаем Gulp
    sass         = null,            // Подключаем Sass препроцессор
    browserSync  = null,         // автоматическая перезагрузка страниц и локальный сервер
    autoprefixer = null,    // автоматическая подстановка префиксов
    cssmin       = null,
    ts           = require("gulp-typescript"),                // typescript
    tsProject    = ts.createProject("tsconfig.json");

gulp.task("typescript", function () {
    return tsProject
        .src()
        .pipe(tsProject())
        .js
        .pipe(gulp.dest("."));
});

    // haml         = require('gulp-haml');
// компилирует sass/scss файлы, добавляя вендорные префиксы
gulp.task('sass', function () {
    return gulp.src('app/scss/*.+(sass|scss)') 
        .pipe(sass())
        .pipe(autoprefixer({
            overrideBrowserslist: ['last 5 versions'],
            cascade: true
        }))
        .pipe(gulp.dest('app/css')) 
        .pipe(browserSync.reload({ stream: true }));
});

gulp.task('browser-sync', function () { // Создаем таск browser-sync
    browserSync({ // Выполняем browserSync
        server: { // Определяем параметры сервера
            baseDir: 'app' // Директория для сервера - app
        },
        notify: false // Отключаем уведомления
    });
});

gulp.task('scripts', function () {
    return gulp.src('app/js/**/*.js')
        .pipe(browserSync.reload({ stream: true }));
});

gulp.task('html', function () {
    return gulp.src('app/**/*.html')
        .pipe(browserSync.reload({ stream: true }));
});

// gulp.task('haml', function () {
//     return gulp.src('app/haml/**/*.haml')
//         .pipe(haml({
//             compiler: 'creationix',
//             compilerOpts: {
//                 minimize: false
//         } }))
//         .pipe(gulp.dest('app/'));
// });

// gulp.task('watch', function () {
//     gulp.watch('app/scss/**/*.+(sass|scss)', gulp.parallel('sass')); // Наблюдение за sass файлами
//     gulp.watch('app/**/*.html', gulp.parallel('html')); // Наблюдение за HTML файлами в корне проекта
//     // gulp.watch('app/haml/**/*.haml', gulp.parallel('haml')); // наблюдение за haml файлами
    // gulp.watch(['app/typescript/**/*.ts'], gulp.parallel('typescript'));
//     gulp.watch(['app/js/*.js', 'app/libs/**/*.js'], gulp.parallel('scripts')); // Наблюдение за главным JS файлом и за библиотеками
// });

gulp.task('watch', function () {
    gulp.watch('**/*.ts', gulp.parallel('typescript'));
});

gulp.task('prebuild', async function () {

    var buildCss = gulp.src(['app/css/*.css'])
        .pipe(cssmin())
        .pipe(gulp.dest('dist/css'));

    var buildFonts = gulp.src('app/fonts/**/*') // Переносим шрифты в продакшен
        .pipe(gulp.dest('dist/fonts'));

    var buildJs = gulp.src('app/js/**/*') // Переносим скрипты в продакшен
        .pipe(gulp.dest('dist/js'));

    var buildImg = gulp.src('app/img/**/*')
        .pipe(gulp.dest('dist/img'));
    
    var buildHtml = gulp.src('app/*.html') // Переносим HTML в продакшен
        .pipe(gulp.dest('dist'));

    var buildLibs = gulp.src('app/libs/**/*')
        .pipe(gulp.dest('dist/libs'));
});

// gulp.task('default', gulp.parallel('sass','browser-sync','watch','prebuild'));
gulp.task('default', gulp.parallel('watch'));
// gulp.task('build', gulp.parallel('prebuild'));