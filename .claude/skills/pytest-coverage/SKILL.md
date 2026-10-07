---
name: pytest-coverage
description: 'Run pytest tests with coverage, discover lines missing coverage, and increase coverage to 100%.'
---

> Source: [github/awesome-copilot](https://github.com/github/awesome-copilot) `skills/pytest-coverage` (commit e263764, MIT — see LICENSE).
> Instructions reworded as recommendations; behavior unchanged. Requires the `pytest-cov` package. `cov_annotate/` is generated output (gitignored).

Recommended goal: tests that cover all lines of code.

Generate a coverage report with:

pytest --cov --cov-report=annotate:cov_annotate

If you are checking for coverage of a specific module, you can specify it like this:

pytest --cov=your_module_name --cov-report=annotate:cov_annotate

You can also specify specific tests to run, for example:

pytest tests/test_your_module.py --cov=your_module_name --cov-report=annotate:cov_annotate

Open the cov_annotate directory to view the annotated source code.
There will be one file per source file. If a file has 100% source coverage, it means all lines are covered by tests, so you do not need to open the file.

For each file that has less than 100% test coverage, find the matching file in cov_annotate and review the file.

If a line starts with a ! (exclamation mark), it means that the line is not covered by tests.
Consider adding tests that cover the missing lines.

Recommended workflow: keep running the tests and improving coverage until all lines are covered, or until a coverage target the user sets is reached.
