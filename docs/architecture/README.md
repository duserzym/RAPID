# RapidPy system map — source and regeneration

`rapidpy.architecture.json` is the typed source of truth for the interactive
system map published at
<https://duserzym.github.io/RAPID/architecture.html>.

It is rendered with [archify](https://github.com/tt-a1i/archify) (MIT), a
standalone Node renderer: no API key, no LLM at render time. Editing the JSON
and re-rendering is the only supported way to change the diagram — do not hand-
edit the generated HTML.

## Generated artifacts

| Artifact | Purpose |
|---|---|
| `docs/site/architecture.html` | Self-contained interactive viewer, deployed by the Pages workflow |
| `docs/site/images/rapidpy-architecture.png` | Static banner used by the README and the site homepage |

## Regenerate

Requires Node 18+ and a checkout of archify.

```bash
git clone --depth 1 https://github.com/tt-a1i/archify.git /tmp/archify
```

Validate the source, then publish the artifact. `--repo-root` lets archify
verify every `sources` path in the JSON against this repository, so a renamed
or deleted module fails the render instead of shipping a stale diagram.

```bash
node /tmp/archify/archify/bin/archify.mjs validate architecture docs/architecture/rapidpy.architecture.json --quality showcase --repo-root . --json
```

```bash
node /tmp/archify/archify/bin/archify.mjs deliver architecture docs/architecture/rapidpy.architecture.json docs/site/architecture.html --quality showcase --repo-root . --json
```

Refresh the static banner from the rendered artifact. `visual-check` writes
screenshots beside the HTML; keep the 2048-wide light one, crop off the partly
visible card row, and delete the rest so they are not deployed.

```bash
node /tmp/archify/archify/bin/archify.mjs visual-check docs/site/architecture.html --json
```

## Rules the source has to satisfy

`--quality showcase` enforces a strict composition budget. Two constraints
drive most of the layout decisions in this file:

- **Canvas width.** Sublabels render at a fixed 9px and every node label,
  sublabel, and boundary label must project to at least 6px at a 930px reader
  width. That caps the viewBox at roughly 1390px, which is why the diagram is
  five columns wide and why each device node names both the instrument and its
  RapidPy adapter instead of splitting them into separate columns.
- **Corridors.** Two hand-routed lines cannot run parallel down the same
  110px gap between columns, and an edge may not pass through an unrelated
  node. The `via` points in the JSON are chosen to keep each corridor to one
  line.

When validation fails it prints the exact fix (`labelDy +24`, `labelAt [x, y]`,
a suggested route change). Apply what it suggests rather than guessing.

## Keeping it honest

`meta.repository.revision` pins the commit the `sources` line links resolve
against. Bump it when the referenced files move, and re-run `deliver` so the
evidence count in the output reflects reality. The deep links only resolve once
that commit is pushed to GitHub.

The diagram describes the software structure. It is **not** a statement about
hardware validation — see
[`docs/rapidpy_transition_readiness_2026-08-29.md`](../rapidpy_transition_readiness_2026-08-29.md)
for what is code-complete versus what still needs physical acceptance.
