const publish = require('mono-pub')
const git = require('@mono-pub/git')
const github = require('@mono-pub/github')
const npm = require('@mono-pub/npm')
const commitAnalyzer = require('@mono-pub/commit-analyzer')

const BREAKING_KEYWORDS = ['BREAKING CHANGE', 'BREAKING-CHANGE', 'BREAKING CHANGES', 'BREAKING-CHANGES']

publish(
    ['package.json'],
    [
        git(),
        github({
            extractCommitsFromSquashed: false,
            releaseNotesOptions: {
                rules: [
                    { breaking: true, section: '⚠️ BREAKING CHANGES' },
                    { type: 'feat', section: '🦕 New features' },
                    { type: 'fix', section: '🐞 Bug fixes' },
                    { type: 'perf', section: '🚀 Performance increases' },
                    { dependency: true, section: '🌐Dependencies' },
                ],
                breakingNoteKeywords: BREAKING_KEYWORDS,
            },
        }),
        commitAnalyzer({
            minorTypes: ['feat'],
            patchTypes: ['perf', 'fix'],
            breakingNoteKeywords: BREAKING_KEYWORDS,
        }),
        npm({ provenance: true, trustedPublishing: true }),
    ]
)