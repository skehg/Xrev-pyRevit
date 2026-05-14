---
name: pyrevit-ironpython27
description: 'Write and review pyRevit tools for IronPython 2.7 with safe Revit API patterns, transaction discipline, and pyRevit-compatible UI. Use when creating or modifying script.py tools, bundle.yaml-backed commands, or Revit element mutation logic.'
argument-hint: 'Task or feature to implement in a pyRevit IronPython 2.7 tool'
user-invocable: true
disable-model-invocation: false
---

# pyRevit IronPython 2.7 Coding Workflow

## Outcome
Produce production-ready pyRevit code that runs on IronPython 2.7, follows repository conventions, and behaves safely in Revit (including Undo-friendly transactions).

## When To Use
- Creating a new `.pushbutton` command with `bundle.yaml` and `script.py`
- Modifying existing pyRevit command logic
- Refactoring Revit API calls for compatibility or reliability
- Adding WPF/pyRevit forms UI to a command
- Reviewing pyRevit code for correctness before release

## Inputs
- Target command folder and files (typically `bundle.yaml`, `script.py`, optional `.xaml`)
- Desired behavior (read-only analysis vs model mutation)
- Revit context constraints (project vs family document, selection requirements)

## Procedure
1. Confirm command structure and scope.
- Locate the target `.pushbutton` folder.
- Verify `bundle.yaml` exists and aligns with command intent.
- Identify whether the script is analysis-only or mutates Revit elements.

2. Establish IronPython 2.7 compatibility baseline.
- Avoid Python 3-only features (f-strings, pathlib-centric assumptions, dataclasses, typing-only runtime dependencies).
- Prefer straightforward Python 2.7-compatible syntax and standard library usage.
- Keep imports and error handling simple and explicit for pyRevit runtime stability.

3. Resolve Revit context up front.
- Use `__revit__` as `UIApplication`; use `.Application` once to access DB application.
- Avoid chaining `.Application.Application`.
- Determine active document type and guard operations accordingly.

4. Choose transaction strategy using decision points.
- If operation is read-only: do not start a transaction.
- If operation mutates elements: wrap writes in `with revit.Transaction("Action Name"):` when possible.
- For renaming or parameter edits, prefer direct native setters inside the transaction first, then fallback patterns only if needed.

5. Implement UI and user flow in pyRevit-safe patterns.
- Prefer `pyrevit.forms` for lightweight prompts and selections.
- If WPF/XAML is used, keep bindings and event handlers simple and deterministic.
- Validate all user inputs before modifying the model.

6. Add robust guardrails and messages.
- Handle empty selection, wrong document type, missing parameters, and API exceptions.
- Return clear user-facing status in pyRevit output.
- Fail early when prerequisites are missing.

7. Validate behavior and quality checks.
- Confirm command executes without syntax/runtime incompatibilities in IronPython 2.7.
- Confirm transaction boundaries are minimal and Undo stack behavior is correct.
- Confirm no unintended document writes in read-only paths.
- Confirm naming/parameter updates apply only to intended elements.

8. Final review against repository conventions.
- Keep pyRevit folder/file layout consistent.
- Preserve existing public behavior unless change is intentional.
- Keep edits focused and avoid unrelated refactors.

## Decision Matrix
- Read-only analysis task:
  - No transaction.
  - Strong null/selection checks.
- Element mutation task:
  - Use `with revit.Transaction("..."):`.
  - Validate targets before write.
- Rename/update task:
  - Try native setter first inside transaction.
  - Use fallback only when API requires it.
- UI-heavy task:
  - Prefer `pyrevit.forms`; use WPF only when necessary.

## Definition Of Done
- Script is IronPython 2.7 compatible.
- Revit writes occur only inside correct transaction scope.
- User input and context checks prevent invalid operations.
- Undo behavior is preserved for mutating commands.
- Command remains consistent with repository structure and patterns.

## Example Invocations
- `/pyrevit-ironpython27 Add a new command that batch renames family types using a safe transaction and preview step.`
- `/pyrevit-ironpython27 Refactor this script.py to remove Python 3 syntax and make it IronPython 2.7 safe.`
- `/pyrevit-ironpython27 Review this pyRevit tool for transaction and Revit API misuse.`
