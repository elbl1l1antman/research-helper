# Legacy Archive

This folder stores deprecated or superseded project assets that should remain available for reference but should not be used by the current launcher-based workflow.

## Rules

- Keep active code in the main project folders.
- Move files here only when a replacement is already committed.
- Preserve enough context to recover the old behavior: original path, retired version, and reason.
- Do not import or execute files from this folder in production workflows.

## Suggested Layout

```text
old/legacy/
  v0.0.x/
    README.md
    ...
```

Each version folder should explain:

- what was retired
- why it was retired
- what replaced it
