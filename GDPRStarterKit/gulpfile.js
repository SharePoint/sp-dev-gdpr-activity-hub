'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

require('./gulpfile-update-manifest.js');

build.initialize(gulp);
