# Contributing

1. Open an issue for a bug or a feature first, unless the change is small. A bug report has a reproduction, the
   expected result and the actual result.
2. Create a branch from `master`, and make one change per pull request.
3. For a bug fix, add a test that fails before the fix. Run `npm run format`, `npm run lint`, `npm run typecheck` and
   `npm test` before you push. CI runs them on Node 22 and 24.
4. Add each breaking change to [UPGRADING.md](UPGRADING.md).
5. Describe in the pull request the problem, a reproduction, the fix, the breaking changes and the tests.

[AGENTS.md](AGENTS.md) describes the code, the commands and the conventions. `npm start` serves a page with a grid on
[http://localhost:8000](http://localhost:8000/).

## Releases

The **Publish** workflow (`.github/workflows/publish.yml`) publishes `@bitbindle/gridwell` to npm. Run it from the
Actions tab, on `master`. The major and minor version come from `package.json`, and the workflow gives the next patch:
with `6.0.0` in `package.json` and `6.0.0` to `6.0.4` on npm, it publishes `6.0.5`. For a new minor or major version,
change `package.json` to `6.1.0` or `7.0.0` in a pull request. The workflow runs the checks and the tests, publishes
with provenance, and pushes the tag `v6.0.5`.

The workflow uses npm trusted publishing: npm trusts this workflow, and no npm token is stored in GitHub. A maintainer
sets it up once on npmjs.com:

1. Create the npm organization `bitbindle`, which owns the scope `@bitbindle`.
2. The trusted publisher is a setting of a package, so publish the first version by hand. Log in with `npm login`, then
   on an up-to-date `master` run `npm install` and `npm publish --access public`. It publishes the version in
   `package.json`, `6.0.0`.
3. In the settings of the package on npmjs.com, add a trusted publisher for GitHub Actions: the organization
   `BitBindle`, the repository `gridwell`, the workflow `publish.yml`, and no environment.
4. In the same settings, set the publishing access to require two-factor authentication and disallow tokens. The
   workflow still publishes, and a stolen token cannot.

Trusted publishing needs npm 11.5.1 or later, which the workflow installs, and a runner hosted by GitHub. The
`repository` in `package.json` must stay `https://github.com/BitBindle/gridwell`, the repository that the
provenance names. If the repository gets a new name, change `repository` and the trusted publisher on npmjs.com before
the next publish.

# Code of Conduct

### Our Pledge

In the interest of fostering an open and welcoming environment, we as
contributors and maintainers pledge to making participation in our project and
our community a harassment-free experience for everyone, regardless of age, body
size, disability, ethnicity, gender identity and expression, level of experience,
nationality, personal appearance, race, religion, or sexual identity and
orientation.

### Our Standards

Examples of behavior that contributes to creating a positive environment
include:

- Using welcoming and inclusive language
- Being respectful of differing viewpoints and experiences
- Gracefully accepting constructive criticism
- Focusing on what is best for the community
- Showing empathy towards other community members

Examples of unacceptable behavior by participants include:

- The use of sexualized language or imagery and unwelcome sexual attention or
  advances
- Trolling, insulting/derogatory comments, and personal or political attacks
- Public or private harassment
- Publishing others' private information, such as a physical or electronic
  address, without explicit permission
- Other conduct which could reasonably be considered inappropriate in a
  professional setting

### Our Responsibilities

Project maintainers are responsible for clarifying the standards of acceptable
behavior and are expected to take appropriate and fair corrective action in
response to any instances of unacceptable behavior.

Project maintainers have the right and responsibility to remove, edit, or
reject comments, commits, code, wiki edits, issues, and other contributions
that are not aligned to this Code of Conduct, or to ban temporarily or
permanently any contributor for other behaviors that they deem inappropriate,
threatening, offensive, or harmful.

### Scope

This Code of Conduct applies both within project spaces and in public spaces
when an individual is representing the project or its community. Examples of
representing a project or community include using an official project e-mail
address, posting via an official social media account, or acting as an appointed
representative at an online or offline event. Representation of a project may be
further defined and clarified by project maintainers.

### Enforcement

Instances of abusive, harassing, or otherwise unacceptable behavior may be
reported by contacting the project team. All
complaints will be reviewed and investigated and will result in a response that
is deemed necessary and appropriate to the circumstances. The project team is
obligated to maintain confidentiality with regard to the reporter of an incident.
Further details of specific enforcement policies may be posted separately.

Project maintainers who do not follow or enforce the Code of Conduct in good
faith may face temporary or permanent repercussions as determined by other
members of the project's leadership.

### Attribution

This Code of Conduct is adapted from the [Contributor Covenant][homepage], version 1.4,
available at [http://contributor-covenant.org/version/1/4][version]

[homepage]: http://contributor-covenant.org
[version]: http://contributor-covenant.org/version/1/4/
