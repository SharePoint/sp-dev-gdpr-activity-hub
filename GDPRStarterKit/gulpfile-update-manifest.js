'use strict';

const build = require('@microsoft/sp-build-web');
const fs = require('fs');

build.task('update-manifest', {
    execute: (config) => {
        return new Promise((resolve, reject) => {
            const cdnPath = config.args['cdnpath'] || "";
            let json = JSON.parse(fs.readFileSync('./config/write-manifests.json'));
            json.cdnBasePath = cdnPath;
            fs.writeFileSync('./config/write-manifests.json', JSON.stringify(json));
            resolve();
        });
    }
});