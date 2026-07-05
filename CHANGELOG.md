# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
The version in `pyproject.toml` (`[project].version`) is the single source of truth.

## [Unreleased]

### Added
- Ruff (lint + formatter) configuration in `ruff.toml` and a `dev` dependency group in `pyproject.toml`.
- `PR check` GitHub Actions workflow running `ruff format --check` and `ruff check` on pull requests.
- `pre-commit` configuration (file-hygiene hooks and Conventional Commits enforcement).
- This changelog.

## [0.1.0]

### Added
- Initial Microsoft Teams bot (Azure Function) automating Databricks Unity Catalog (DPS) access requests.
