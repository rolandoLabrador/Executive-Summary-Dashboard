# Workspace Rules & Preferences

## 1. Review-First Workflow
- **No Direct/Silent Edits**: Always present proposed code changes, file paths, and diffs to the user first.
- **Wait for Explicit Approval**: Do not modify or create project source files until the user reviews and confirms the plan.

## 2. Educational & Clean Code Standards
- **Explain Rationale**: For every change, provide a clear explanation of *why* this specific approach/pattern was chosen and why it represents the cleanest, most maintainable solution.
- **Key Principles**:
  - **Single Source of Truth (DRY)**: Centralize thresholds, constants, and shared logic.
  - **Separation of Concerns**: Keep data transformation/business logic separate from UI/presentation logic (e.g., Excel styling vs. data calculation).
  - **Defensive & Type-Safe Code**: Use TypeScript types, nullish coalescing (`??`), and bounded ranges to prevent runtime crashes and style bleeding.
  - **Dynamic Layouts**: Avoid hardcoded magic numbers; calculate coordinates based on dynamic data bounds.
  - **Automated Verification**: Always verify changes against the test suite (`npm test`) and build system (`npm run build`).
