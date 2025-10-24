# Minimal Prompt v3 (class-level `examples`)

You are an expert API documentation structurer.
You will receive a **class documentation** snippet in Markdown and must convert it into **strict JSON** following the schema and rules below.
**Output JSON only** (no prose, no Markdown fences).

## Output Rules

* Output **JSON only**, UTF-8, no extra text or fields.
* If data is missing, use `null` or `[]` according to the schema.
* **Do not invent examples**: only populate `examples` where the input actually provides them; otherwise use `[]`.
* Merge method overloads under the same method `name` using a `signatures` array.
* Infer `method.kind` when clear; if uncertain, set `null`.

  * Allowed values: `read | write | create | delete | load | configure | serialize | track | untrack | null`.

## Target JSON Schema

```json
{
  "class": {
    "name": "string",
    "package": "string|null",
    "extends": ["string"],
    "api_set": { "name": "string|null", "status": "string|null" },
    "description": "string|null",
    "examples": [
      {
        "description": "string|null",
        "usage_code": "string|null",
        "output_code": "string|null"
      }
    ]
  },
  "properties": [
    {
      "name": "string",
      "type": "string|null",
      "description": "string|null",
      "since": "string|null",
      "examples": [
        {
          "description": "string|null",
          "usage_code": "string|null",
          "output_code": "string|null"
        }
      ]
    }
  ],
  "methods": [
    {
      "name": "string",
      "kind": "read|write|create|delete|load|configure|serialize|track|untrack|null",
      "description": "string|null",
      "signatures": [
        {
          "params": [
            { "name": "string", "type": "string|null", "required": true, "description": "string|null" }
          ],
          "returns": { "type": "string|null", "description": "string|null" }
        }
      ],
      "examples": [
        {
          "description": "string|null",
          "usage_code": "string|null",
          "output_code": "string|null"
        }
      ]
    }
  ],
  "source": {
    "urls": ["string"]
  }
}
```

## Parsing Guidance

1. **Class block**: Extract `name`, `package`, `extends`, `api_set` (from “API set/Remarks”), and `description` from the title and overview sections.

   * Put class-level usage demos under `class.examples`.
2. **Properties**: From “Properties / Property details”, extract `name`, `type` (prefer TypeScript declarations), `description`, and `since` (from API set/remarks).

   * If any property includes example code/snippets, place them under that property’s `examples`.
3. **Methods**: From “Methods / Method details”, group overloads of the same method under one `name` with multiple `signatures` (each with `params` and `returns`).

   * Determine `kind` by semantics (`read`, `create`, `delete`, etc.); if unclear, set `null`.
   * Place method usage examples under that method’s `examples`.
4. **Examples policy**: Do **not** generate examples if the input doesn’t provide them—leave `examples: []` at that level.
5. **Links**: Collect authoritative doc links into `source.urls`.

## Input

Paste the entire class Markdown here, unmodified, between delimiters:

<<<
[PASTE CLASS MARKDOWN HERE]
>>>

**Return only the JSON per the schema above.**
