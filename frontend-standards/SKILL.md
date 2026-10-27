---
name: frontend-standards
version: 1.1.0
description: >
  Enforce team coding standards across frontend projects. Triggers when writing
  or editing component code, styles, or utility functions — in Vue, React, Svelte,
  or any frontend framework. Also triggers when the user mentions coding standards,
  style conventions, code review, formatting rules, or project conventions. Use
  even for simple style edits if they touch a <style> block or function declaration.
---

## Conventions

This skill collects all agreed-upon coding standards. Each convention below
explains **what**, **why**, and **when** to apply it.

### 1. SCSS: Nested selectors

**Rule:** All SCSS in component `<style>` blocks must use nested selectors (`&` and nested rules). Flat top-level class selectors are forbidden. Inside a block, write variants as `&__element` / `&--modifier` / `&:state` — never nest a full class name. Nesting depth must not exceed 3 levels.

**Why:** Nested SCSS mirrors the DOM tree, so readers see parent/child relationships without cross-referencing the template. Using `&` keeps that readability **without** inflating specificity: `&__title` compiles to `.card__title` (one class), while nesting `.card__title` inside `.card` compiles to `.card .card__title` (descendant selector) — harder to override later and coupled to DOM shape.

**How to apply:**

```scss
// Correct — `&` keeps the compiled selector a single class
.card {
  padding: 20px;
  background: #fff;

  &__title {
    font-weight: 700;

    &--muted {
      color: #888;              // → .card__title--muted
    }
  }

  &--active {
    border-color: #000;         // → .card--active
  }

  &:hover {
    box-shadow: 0 2px 8px rgb(0 0 0 / 10%);  // → .card:hover
  }
}

// Forbidden — flat top-level selectors
.card { }
.card__title { }
.card__title--muted { }
.card:hover { }

// Also forbidden — nesting a full class name
.card {
  .card__title { }              // → .card .card__title (needless specificity)
}
```

**Nesting depth:** `a { .b { .c { } } }` is the limit — 3 levels. Deeper structure means a child component, not deeper nesting. Use `@at-root` when a rule must escape its parent.

**Exceptions:** `@keyframes` blocks, global utility classes in shared style files, `:root` / `html` variable declarations.

**Deep selectors (Vue scoped):** Use `:deep()` inside the nesting:
```scss
.form-panel {
  :deep(.wd-input__inner) { border-radius: 4px; }
}
```

### 2. Functions: Arrow functions

**Rule:** All function declarations should use arrow syntax (`const fn = (...) => { ... }`). Use `function` keyword only when you need a dynamic `this` (event listeners, class methods, lifecycle hooks where `this` matters).

**Why:** Arrow functions are shorter, more consistent, and avoid the common `this` confusion. They also read more naturally — `const formatPrice = (price) => ...` states the function name first, before the "noise" of `function`.

**How to apply:**

```ts
// Correct
const formatPrice = (price: number): string => `¥${price.toFixed(2)}`

const handleSubmit = async (data: FormData) => {
  await submit(data)
  resetForm()
}

// Acceptable (function needed for `this` or hoisting)
function handleClick(this: HTMLButtonElement, event: MouseEvent) {
  this.classList.toggle('active')
}

// Forbidden
function formatPrice(price: number): string {
  return `¥${price.toFixed(2)}`
}
```

**Exceptions:**
- Functions that need a dynamic `this` (class methods, DOM event handlers using `this`)
- Named function expressions that benefit from hoisting (e.g., early returns before definition)
- Generator functions (`function*`)

### 3. Documentation: Sync on code change

**Rule:** When you modify source files (`.ts`, `.vue`, `.js`, `.scss`, `.css`), always check whether related documentation needs updating before the session ends. Use the checklist below for any repo whose docs reference the changed code.

**Why:** Docs rot is the #1 source of onboarding friction. A correct codebase with wrong docs is worse than no docs at all — it creates false confidence and wastes debugging time. Every code-only change widens the gap between what the code *does* and what the team *thinks* it does.

**How to apply:**

1. Before ending a coding session, scan the git diff for changed files
2. For each changed module/component, check if there are related docs:
   - `README.md` — architecture overview, setup steps, environment variables
   - `CHANGELOG.md` — user-facing changes
   - `docs/` — API docs, component stories, usage guides
   - `*.stories.*` — Storybook stories
3. If docs reference the changed behavior, update them in the same session
4. If no docs exist but the change is significant, consider creating a doc entry

**Trigger checklist:**

```
Code change touches        →  Check these docs
─────────────────────────────────────────────────
API/endpoint signature      →  API docs, README
Component props/slots       →  Storybook, usage guide
Config/env vars             →  README setup section
New dependency              →  README prerequisites
Breaking behavior change    →  CHANGELOG, migration guide
SCSS/tokens change          →  Design system docs
```

**Exceptions:**
- Trivial refactors that don't change behavior or API surface
- Internal implementation details not referenced by any doc
- One-line bug fixes that restore intended behavior (not change it)

<!-- Add new conventions below this line -->
