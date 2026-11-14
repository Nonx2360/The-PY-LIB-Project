# THE PY-LIB — Feature Expansion Implementation Plan

> **Goal:** Add four practical school-library workflows to THE PY-LIB without weakening its existing data-integrity, security, offline-first, and cross-interface architecture.
>
> **Priority order:** P0 Academic Year Migration → P0 Inventory Audit → P1 Self-Service Kiosk → P1 Book Reservation.
>
> **Audience:** Coding agents and developers implementing this plan in the existing repository.

---

## 0. Instructions for the Coding Agent

Before changing code:

1. Read `README.md`, `docs/SECURITY.md`, and `docs/DESIGN.md`.
2. Inspect the actual database schema and all current migrations in `db/schema.py`.
3. Inspect the relevant repositories, services, Flask routes, templates, desktop views, WebSocket events, and tests.
4. Map existing member fields, book fields, loan rules, statuses, QR resolution, session handling, CSRF protections, and backup behavior. **Do not assume field names or invent a second source of truth.**
5. Check whether any part of the requested features already exists. Extend it rather than duplicating it.
6. Write a short implementation proposal with the exact files and migration changes before coding.
7. Preserve current behavior unless a change is explicitly required. Do not rewrite the application architecture or replace its libraries.

### Existing architecture to preserve

THE PY-LIB currently has:

- A CustomTkinter desktop administration application.
- A Flask administration web interface.
- A separate student Member Portal.
- A shared SQLite database and shared repositories/services.
- Atomic borrow/return operations in the repository layer.
- QR identity resolution, including encrypted member QR and expiring rotating tickets.
- Separate admin/member authentication contexts, bcrypt PIN/password handling, CSRF protection, throttling, and role guards.
- WebSocket-based live updates.
- SQLite backup/restore and schema migrations through `PRAGMA user_version`.
- Vendored frontend assets and local fonts for offline use.

**Architectural rule:** Desktop, admin web, kiosk, and Member Portal must call the same service/repository business logic. Do not implement separate borrowing, reservation, migration, or inventory rules in the UI layer.

---

## 1. Scope and Priority

| Priority | Feature | Main outcome |
|---|---|---|
| P0 | Academic Year Migration | Safely promote students to the next academic year/class while preserving history |
| P0 | Inventory Audit | Reconcile physical books against database records and identify missing/unexpected books |
| P1 | Self-Service Kiosk | Let students perform permitted self-service actions through a restricted kiosk interface |
| P1 | Book Reservation | Let students queue for currently unavailable books and receive an offer when one is returned |

### Global non-goals

- Do not introduce a cloud dependency or paid API.
- Do not replace SQLite, Flask, CustomTkinter, or the current frontend stack.
- Do not expose admin functions or database internals to students.
- Do not store plaintext PINs, passwords, activation codes, or reusable QR secrets.
- Do not silently alter historical loans, returns, access logs, or audit records.
- Do not auto-delete records during migration or inventory reconciliation.
- Do not make the kiosk a second admin interface.

---

# 2. P0 — Academic Year Migration

## 2.1 Purpose

Provide a controlled workflow for moving student records into a new academic year. This should reduce repetitive manual editing while protecting historical borrowing records and students who leave the school.

## 2.2 User stories

- As an administrator, I can preview how student records will change before applying a promotion.
- I can promote eligible students to the next class/grade in bulk.
- I can keep selected students unchanged, mark them as graduated/left, or handle them manually.
- I can export a preview and a result report.
- I can recover safely if the operation fails.

## 2.3 Functional requirements

### A. Academic year configuration

- Inspect the existing member schema and settings first.
- Define an explicit academic-year representation appropriate to the project's current conventions.
- Support a configurable promotion mapping, rather than hardcoding assumptions about every school's class structure.
- Support a dry-run preview before any write occurs.
- Make the target academic year explicit and prevent accidental duplicate execution for the same source/target year unless an administrator intentionally starts a new run.

### B. Promotion preview

Show a review table containing, at minimum:

- Current student identity fields already present in the database.
- Current class/grade.
- Proposed class/grade.
- Proposed status/action: promote, remain unchanged, graduated/left, or manual review.
- A reason when a record cannot be promoted automatically.

Allow administrators to:

- Search/filter the preview.
- Include/exclude individual records.
- Apply a supported action to a selected group.
- Export the preview to CSV or Excel using existing export utilities where suitable.

### C. Promotion execution

- Require an explicit confirmation that names the source and target academic years and affected record count.
- Create a verified backup before applying changes, using the existing backup service.
- Apply all eligible changes in a single controlled database transaction where feasible.
- Record a migration batch/run record with actor, timestamps, source year, target year, counts, and outcome.
- Make the operation idempotent or detect duplicate runs.
- Preserve stable member identifiers so historical loans, access records, and other relationships remain intact.
- Do not overwrite historical records with the student's new class/grade.
- If the current schema has no safe way to distinguish current student placement from historical placement, propose a schema design before implementation. Do not repurpose a field that historical records depend on.

### D. Exceptions and rollback

- Graduated/left students must not be deleted.
- Students with incomplete or ambiguous data must be flagged for manual review.
- A failed transaction must leave no partially promoted set.
- Provide a clear completion report: promoted, unchanged, graduated/left, skipped, and failed.
- Document the restore procedure. Do not build an unsafe "undo" that guesses old values.

## 2.4 Suggested data design

The coding agent must validate this against the real schema before implementation.

Potential entities:

- `academic_years`: academic-year identifier, display label, start/end dates, active state.
- `student_placements` (or equivalent): stable member reference, academic-year reference, class/grade/room/number fields as applicable.
- `academic_year_migrations`: run metadata and status.
- `academic_year_migration_items`: per-member before/after snapshot and action outcome.

If the current member table already has a suitable current-placement model, reuse it. Avoid unnecessary tables.

## 2.5 UI and integration

- Add an admin-only Academic Year Migration page to the Flask web UI.
- Reuse existing design tokens, table patterns, confirmation dialogs, CSRF protection, and role guards.
- Consider a desktop entry point only if it fits the existing navigation patterns; do not create two independent implementations.
- Emit a real-time update after a successful migration so open admin views refresh.
- Update documentation and operational guidance.

## 2.6 Acceptance criteria

- [ ] Admin can configure/select a source and target academic year.
- [ ] Preview is read-only and makes no database changes.
- [ ] Promotion mapping can be reviewed and adjusted before execution.
- [ ] Existing member IDs and historical loan/access records remain intact.
- [ ] Graduated/left students are retained and cannot accidentally be promoted.
- [ ] Duplicate execution is detected or safely prevented.
- [ ] Backup verification occurs before mutation.
- [ ] A failure does not leave a partial migration.
- [ ] A result report is available after completion.
- [ ] Unauthorized users cannot view or execute the workflow.
- [ ] Tests cover normal, skipped, duplicate, rollback, and historical-data preservation cases.

---

# 3. P0 — Inventory Audit

## 3.1 Purpose

Help library staff physically verify the collection and identify discrepancies between books on shelves and the database.

The audit is a reconciliation workflow, **not** an automatic mass-edit or deletion feature.

## 3.2 User stories

- As library staff, I can start an inventory audit session.
- I can scan book QR codes while walking through shelves.
- I can see which books have been checked, are still expected, or were scanned unexpectedly.
- I can pause and resume an audit.
- I can review discrepancies and decide what to do with each one.

## 3.3 Functional requirements

### A. Audit session lifecycle

- Create an audit session with a name/reference, optional location/section, creator, creation time, and status.
- Supported states: draft/in progress, paused, completed, and cancelled.
- Only authorized staff can create, edit, resume, complete, or cancel sessions.
- Preserve completed audit sessions as historical records.
- Prevent accidental duplicate active sessions for the same scope if the existing workflow requires that constraint.

### B. Scan and reconciliation

- Reuse the existing QR decoding and book/member identity resolution patterns where appropriate.
- A scan must resolve to a real book record; never trust client-supplied book details.
- Record each unique book at most once per audit session for the main count, while retaining useful duplicate-scan feedback or scan events if needed.
- Show immediate feedback:
  - Found and expected.
  - Already scanned in this session.
  - Not in the database / unrecognized QR.
  - Retired or otherwise not eligible for the selected audit.
  - Unexpected location/section, if book-location data exists.
- Do not treat a currently borrowed book as physically missing without an explicit audit policy. Show its loan state separately.
- Support manual lookup/addition only with an auditable staff action.

### C. Reconciliation report

At any time, show:

- Expected books in the selected scope.
- Unique books scanned.
- Not-yet-scanned books.
- Unexpected/unrecognized scans.
- Retired books encountered.
- Books currently on loan, separated from shelf-expected items.
- Progress percentage with a clear denominator.

At completion, generate a report that can be viewed and exported using existing report/export patterns.

### D. Discrepancy resolution

For every discrepancy, staff must explicitly choose an action, such as:

- Confirm book found.
- Mark as not found after physical re-check.
- Flag for investigation.
- Correct the book's location if location tracking exists.
- Add a note.

Do not automatically mark books lost, delete books, or change their availability just because they were not scanned.

## 3.4 Suggested data design

Validate against the current schema and existing audit/log patterns.

Potential entities:

- `inventory_audits`: audit metadata, scope, status, creator, timestamps.
- `inventory_audit_scans`: audit reference, book reference, scanned time, scanner/operator, result.
- `inventory_audit_resolutions`: discrepancy, selected resolution, note, actor, timestamp.

Use foreign keys and indexes appropriate to the existing SQLite conventions. If a book location/shelf field does not exist, make location support a separate optional enhancement rather than assuming it exists.

## 3.5 UI and integration

- Add an admin-only Inventory Audit page.
- Provide a scan-first interface that works with the current QR scanning approach.
- Support keyboard/manual search as a fallback when camera access is unavailable.
- Make the interface usable on a tablet or laptop.
- Use existing SweetAlert/toast/table components and design tokens.
- Push updates to relevant admin clients through the existing WebSocket mechanism only after committed changes.

## 3.6 Acceptance criteria

- [ ] Staff can create, pause, resume, complete, and cancel an audit.
- [ ] A book scan is resolved server-side against the database.
- [ ] Duplicate scans do not inflate the unique count.
- [ ] Unknown QR values do not create fake book records.
- [ ] On-loan books are represented separately from shelf-missing books.
- [ ] Unscanned books are not automatically marked lost or deleted.
- [ ] Staff can review and resolve discrepancies with an audit trail.
- [ ] Completed sessions remain available for historical review.
- [ ] Reports show counts and a clear progress denominator.
- [ ] Tests cover duplicates, unknown QR, borrowed books, permissions, and resolution history.

---

# 4. P1 — Self-Service Kiosk

## 4.1 Purpose

Provide a restricted, student-facing kiosk so students can complete approved library actions without accessing the administration interface.

The first release should prioritize a safe, reliable self-service flow over a large feature set.

## 4.2 Recommended first-release scope

1. Identify the student using the existing member QR identity flow.
2. Confirm identity server-side.
3. Show only the minimum information needed for the transaction.
4. Scan a book QR code.
5. Display a clear transaction summary.
6. Submit the request through the existing atomic borrowing/returning services.
7. Show a success/failure receipt.
8. Clear the session and return to the welcome screen automatically.

If the current loan policy or QR protocol does not support a safe student-initiated return, keep return as staff-assisted in the first release. Do not weaken the existing authorization model to force the feature.

## 4.3 Functional requirements

### A. Kiosk mode

- Dedicated kiosk routes/templates, separate from admin routes.
- Full-screen, touch-friendly layout with large controls.
- No admin navigation, admin data tables, settings, or arbitrary URL access from the kiosk UI.
- Clear idle timeout and automatic reset to the welcome screen.
- Provide a visible cancel/reset action.
- Do not rely on hiding buttons as authorization; enforce permissions on the server.

### B. Identity and session safety

- Reuse the existing rotating QR/member identity resolver where compatible.
- Validate the QR on the server; never accept a member ID as proof of identity.
- Use a dedicated short-lived kiosk session/context.
- Prevent kiosk identity from becoming an admin session or reusing admin privileges.
- Apply rate limiting and generic error responses to reduce member enumeration.
- Do not display PINs, activation codes, internal IDs, or other students' data.
- Clear sensitive client-side state after completion, cancellation, and timeout.

### C. Transaction flow

- Use the same business rules and repository/service methods as staff-assisted borrowing.
- Enforce loan limits, book availability, member eligibility, due-date policy, and atomic concurrency protection.
- Prevent duplicate submissions caused by double taps or slow network responses.
- Return a transaction result without exposing stack traces or internal database details.
- If the current design requires staff approval, make that explicit rather than bypassing it.

### D. Privacy and physical operation

- Do not show a previous student's information after a transaction.
- Automatically reset after inactivity.
- Avoid printing or displaying unnecessary personal information.
- Provide a staff-only way to exit kiosk mode or perform maintenance.
- Document recommended browser/device setup and camera permission requirements.
- Treat the kiosk as a shared public device.

## 4.4 UI and integration

- Add a dedicated kiosk entry point and templates.
- Reuse the existing QR scanning JavaScript and backend resolver where safe.
- Reuse existing design tokens and local fonts.
- Ensure the kiosk remains usable on common laptop/tablet screen sizes.
- Add clear Thai-language instructions and accessible error states.
- Do not require external CDNs or paid services.

## 4.5 Acceptance criteria

- [ ] Student can identify themselves using the supported secure QR flow.
- [ ] Kiosk cannot access admin routes or admin-only APIs.
- [ ] Kiosk uses the shared atomic borrowing service.
- [ ] Loan policy is enforced identically to staff-assisted transactions.
- [ ] Duplicate requests cannot create duplicate loans.
- [ ] Idle timeout and explicit reset clear the prior student's state.
- [ ] Error messages do not leak whether arbitrary member IDs exist.
- [ ] No sensitive student information remains visible after reset.
- [ ] Camera-unavailable/manual fallback behavior is defined.
- [ ] Tests cover permissions, timeout/reset, rate limits, and concurrent borrowing.

---

# 5. P1 — Book Reservation

## 5.1 Purpose

Allow a student to join a queue for a book that is currently unavailable, then offer the book to the next eligible student when it becomes available.

## 5.2 User stories

- As a student, I can reserve an unavailable book from the Member Portal.
- I can see my active reservations and queue status.
- I can cancel a reservation.
- When a book becomes available for me, I can see the offer and its expiry.
- As an administrator, I can review reservations and resolve exceptional cases.

## 5.3 Functional requirements

### A. Reservation eligibility

- Allow reservations only for books that are eligible under the library's policy.
- Prevent a student from reserving the same book multiple times while they already have an active reservation.
- Define whether a student who already has the book on loan can reserve another copy; follow the current loan policy.
- Define reservation limits per student through configuration if appropriate.
- Do not allow reservations for retired, deleted, or otherwise unavailable-for-circulation books.

### B. Queue behavior

- Use a deterministic FIFO queue based on server-created timestamps plus a stable tie-breaker.
- Queue position must be computed from current active reservations; do not trust a position supplied by the browser.
- Use a transaction when assigning an available copy to the next eligible reservation.
- Handle multiple copies of the same title/book grouping only if the current data model can reliably represent that relationship. Otherwise, scope the first version to the existing book-record model.
- Skip cancelled, expired, or ineligible reservations safely.
- Avoid race conditions when multiple returns or staff actions make copies available concurrently.

### C. Offer and expiry

- When a relevant book becomes available, create an offer for the next eligible student.
- Offer expiry must be configurable and clearly displayed.
- Student can accept/claim the offer or allow it to expire.
- On expiry or cancellation, the next eligible reservation is offered the book.
- Define the exact state transitions and ensure they are enforced server-side.
- Do not silently create a normal loan until the student completes the required claim/checkout step.
- If the current school process requires staff confirmation, include an admin confirmation step.

### D. Notifications

- Initial release: show reservation state and offer notifications inside the Member Portal.
- Reuse existing member WebSocket updates where appropriate.
- Do not add email, SMS, LINE, or third-party notification dependencies in this scope.
- If live updates are unavailable, the page must still show correct state on refresh.

### E. Cancellation and administration

- Students can cancel their own active reservations.
- Admins can inspect, cancel, or resolve reservations with a recorded reason.
- Keep historical reservation records; do not hard-delete completed/cancelled/expired rows.
- Record who performed administrative actions and when.

## 5.4 Suggested state model

Use explicit states and enforce valid transitions. A possible starting model:

- `queued`
- `offered`
- `claimed` (only if the workflow needs a claim stage)
- `fulfilled`
- `cancelled`
- `expired`

The coding agent must finalize the state model after reviewing the actual loan and book-availability rules. Avoid ambiguous combinations of booleans.

## 5.5 Suggested data design

Validate against the actual book/member/loan schema first.

Potential entity:

- `book_reservations`: reservation ID, member reference, book or title/copy reference, state, queue timestamp, offer timestamp, expiry timestamp, claim/fulfilment timestamp, cancellation reason, and audit timestamps.

Add appropriate uniqueness constraints and indexes for active queue lookups. Use foreign keys and the project's established migration conventions.

## 5.6 UI and integration

Member Portal:

- Reserve action only when eligible.
- Active reservation list with current state and queue position.
- Offer card with expiry time and claim action.
- Cancel action with confirmation.
- Empty/loading/error states.

Admin:

- Reservation queue overview.
- Filters by state/book/member.
- Manual resolution actions with confirmation and audit history.

Backend:

- Shared reservation service.
- Transaction-safe queue assignment.
- WebSocket event after committed state changes.
- CSRF and existing session/role protections on all state-changing requests.

## 5.7 Acceptance criteria

- [ ] Eligible student can reserve an unavailable book.
- [ ] Duplicate active reservation is prevented.
- [ ] Queue ordering is deterministic and server-controlled.
- [ ] Student can view and cancel only their own reservations.
- [ ] Admin can review and resolve reservations.
- [ ] A returned/available book is offered to the next eligible student.
- [ ] Offer expiry and queue advancement work correctly.
- [ ] Concurrent availability events cannot offer the same copy to multiple students.
- [ ] Reservation does not create a loan until the defined checkout/claim step.
- [ ] Historical reservation records are retained.
- [ ] Tests cover state transitions, queue order, expiry, cancellation, permissions, and concurrency.

---

# 6. Shared Engineering Requirements

## 6.1 Database migrations

- Follow the existing `PRAGMA user_version` migration approach.
- Inspect the current latest schema version before choosing the next version.
- Make migrations forward-only, deterministic, and safe for existing databases.
- Use transactions where supported.
- Add indexes and foreign keys consistent with current conventions.
- Test migrations from a database at the current production schema version.
- Do not edit old migration behavior in a way that breaks existing installations.
- Ensure a fresh database and an upgraded database produce equivalent schema.

## 6.2 Data integrity

- Keep business rules in services/repositories, not templates or JavaScript.
- Use transactions for multi-row state changes.
- Preserve stable member and book identifiers.
- Keep historical loans, returns, access logs, audit sessions, and reservation history.
- Use uniqueness constraints for invariants that must hold under concurrency.
- Never rely on a prior UI check as the only protection against duplicate or conflicting writes.

## 6.3 Security

- Reuse the current admin/member session separation and role guards.
- Apply CSRF protection to every state-changing web request.
- Apply rate limiting to public/shared-device entry points.
- Validate all identifiers and state transitions server-side.
- Never log PINs, passwords, activation codes, QR payload secrets, or session cookies.
- Avoid exposing raw exception details to students or kiosk users.
- Keep kiosk routes isolated from admin routes.
- Record privileged actions with actor and timestamp.
- Review `docs/SECURITY.md` whenever authentication, authorization, sessions, QR handling, or sensitive data flows change.

## 6.4 Real-time updates

- Use the existing WebSocket infrastructure instead of introducing another event system.
- Publish events only after the database transaction commits.
- Treat WebSocket messages as refresh hints, not as authoritative business data.
- The UI must recover by fetching current server state after reconnect or refresh.

## 6.5 UI and accessibility

- Follow `docs/DESIGN.md` and the current visual system.
- Reuse the existing local fonts and vendored libraries.
- Maintain Thai-language labels and clear date formatting.
- Include loading, empty, success, validation, and failure states.
- Support keyboard operation as well as touch where relevant.
- Avoid unnecessary animation and avoid introducing a generic AI-generated dashboard aesthetic.

---

# 7. Testing Strategy

The repository currently documents a substantial unittest suite and CI checks. Extend the existing test style and helpers rather than introducing a parallel test framework without a clear reason.

For every feature, include:

### Unit tests
- Service-level business rules.
- Valid and invalid state transitions.
- Boundary conditions and duplicate operations.

### Repository/database tests
- Transaction rollback.
- Foreign-key behavior.
- Uniqueness constraints.
- Migration from the current schema.
- Historical-data preservation.

### Web/security tests
- Unauthenticated access.
- Wrong-role access.
- CSRF rejection.
- Rate limiting where applicable.
- Cross-member data isolation.
- No sensitive information in responses.

### Integration/regression tests
- Existing borrowing and returning behavior remains unchanged.
- Existing QR versions continue to resolve as before.
- Admin and Member Portal sessions remain isolated.
- WebSocket notifications occur only after committed changes.
- Existing backup/restore and CI checks continue to pass.

### Required validation before each feature is considered complete

Run the repository's documented commands from the project's supported Python environment:

```powershell
python -m unittest discover -s tests
ruff check .
```

Also run the relevant migration, security, and integration tests. Do not claim success if a command was not actually run.

---

# 8. Recommended Implementation Sequence

Implement one feature at a time. Do not start the next feature until the current one has passed its acceptance criteria and regression suite.

## Phase 0 — Repository discovery and design

- [ ] Read README, SECURITY, and DESIGN docs.
- [ ] Inspect current schema, migration version, repositories, services, routes, templates, and tests.
- [ ] Identify existing functionality that can be reused.
- [ ] Write a concise technical design for all four features.
- [ ] Identify schema assumptions and open policy questions.

## Phase 1 — Academic Year Migration (P0)

- [ ] Finalize academic-year and student-placement model.
- [ ] Add migration and repository support.
- [ ] Implement preview/dry-run service.
- [ ] Implement transaction-safe promotion execution.
- [ ] Add admin UI and result report.
- [ ] Add audit trail and backup preflight.
- [ ] Add tests and update docs.

## Phase 2 — Inventory Audit (P0)

- [ ] Finalize audit scope and treatment of borrowed books.
- [ ] Add audit-session schema and repositories.
- [ ] Implement scan/reconciliation service.
- [ ] Add scan-first admin UI.
- [ ] Add discrepancy resolution and report export.
- [ ] Add tests and update docs.

## Phase 3 — Self-Service Kiosk (P1)

- [ ] Finalize permitted student actions and any staff-approval requirement.
- [ ] Define isolated kiosk session and route permissions.
- [ ] Implement kiosk service/controller using existing transaction logic.
- [ ] Add kiosk UI, scan flow, timeout, and reset.
- [ ] Test shared-device privacy and transaction safety.
- [ ] Document device/browser deployment.

## Phase 4 — Book Reservation (P1)

- [ ] Finalize reservation policy and queue/offer state machine.
- [ ] Add schema, indexes, repositories, and reservation service.
- [ ] Integrate with book-return/availability workflow.
- [ ] Add Member Portal reservation UI.
- [ ] Add admin queue management.
- [ ] Add expiry processing and real-time refresh events.
- [ ] Test concurrency and state transitions.
- [ ] Update docs.

## Phase 5 — Final regression and release

- [ ] Run the complete test suite and lint.
- [ ] Verify fresh-install and existing-database migration paths.
- [ ] Verify backup/restore with the new schema.
- [ ] Manually test desktop, admin web, Member Portal, and kiosk flows.
- [ ] Review permissions and sensitive-data exposure.
- [ ] Update README feature list, architecture, screenshots, and roadmap.
- [ ] Prepare release notes with known limitations.

---

# 9. Definition of Done

A feature is complete only when:

1. Its acceptance criteria are met.
2. Its business rules live in the shared service/repository layer.
3. Database migrations work for both fresh and existing installations.
4. Tests cover normal behavior, edge cases, permissions, and data integrity.
5. Existing features and security controls continue to work.
6. UI follows the existing design system and supports Thai text.
7. Documentation explains configuration, usage, limitations, and recovery.
8. The coding agent reports exactly which tests and commands it ran, including failures.

## 10. Final Agent Instructions

Work incrementally and keep each phase reviewable. Start with **Phase 0 only**: inspect the repository and return a concise technical design, exact files likely to change, schema proposal, risks, and unresolved policy questions. Do not implement all four features in one pass.

After approval, implement **Academic Year Migration first**, then proceed in the priority order above. Preserve the existing architecture and avoid unrelated refactors.
