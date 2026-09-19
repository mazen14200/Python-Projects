# Master Plan — Building a Script to Turn Word Lists into One Combined Image

## 1. Project Goal

Create a single Python script that receives a variable number of English words from the Windows CMD in this exact input style:

```bash
python surreal_compose.py '{"desk","cat","car"}'
```

The script must:

1. Read the words from the command line.
2. Search the Internet for an appropriate image representing each word.
3. Download the selected image temporarily.
4. Remove its background automatically using `rembg`.
5. Create one final artistic composition containing all successfully processed objects.
6. Use a **1024 × 1024** canvas.
7. Use a **light beige / light warm background**.
8. Arrange the objects in an **artistic random composition**, rather than a rigid grid.
9. Produce exactly one final combined image.
10. Delete temporary downloaded and background-removed images after the final image is created.

The design must work with a flexible number of words, including two words, ten words, or more.

---

## 2. Confirmed Input Contract

### 2.1 Command-Line Input

The script accepts one command-line argument containing a Python-set-like string:

```text
{"desk","cat","car"}
```

Example:

```bash
python surreal_compose.py '{"desk","cat","car"}'
```

Another valid example:

```bash
python surreal_compose.py '{"book","dog"}'
```

And a larger example:

```bash
python surreal_compose.py '{"desk","cat","car","book","tree","chair","lamp","phone","clock","flower"}'
```

### 2.2 Input Requirements

- The number of words is not fixed.
- The script must not be hard-coded for exactly three words.
- At minimum, the intended use case is multiple words/objects.
- The input syntax must remain compatible with the agreed set format:
  `{"word1","word2","word3"}`.
- Invalid input should be handled cleanly with a useful error message.

### 2.3 Interpretation

Each word represents an object/concept for which the script should find a suitable image.

For example:

```text
desk → desk image
cat  → cat image
car  → car image
```

The goal is not to create three separate final images. The goal is one final composition containing all selected objects.

---

## 3. End-to-End Processing Flow

The complete pipeline is:

```text
CMD Input
   ↓
Parse Word Set
   ↓
Validate Words
   ↓
Create Temporary Workspace
   ↓
Search Internet for Each Word
   ↓
Download One Suitable Image per Word
   ↓
Remove Background
   ↓
Validate Transparent PNG
   ↓
Resize / Rotate / Position
   ↓
Compose All Objects on 1024×1024 Canvas
   ↓
Save final_composition.png
   ↓
Delete Temporary Files
   ↓
Finish
```

---

# 4. Phase 1 — Parse and Validate Input

## 4.1 Read `sys.argv`

The script reads the first command-line argument.

Expected:

```python
sys.argv[1]
```

Example raw value:

```text
{"desk","cat","car"}
```

## 4.2 Parse Safely

The input should be interpreted as a Python literal rather than executed as arbitrary code.

The implementation should use:

```python
ast.literal_eval(...)
```

and must not use unrestricted `eval(...)`.

## 4.3 Normalize the Words

After parsing:

- Convert the result to a sequence of words.
- Remove accidental surrounding whitespace.
- Ignore empty values.
- Preserve the actual user-supplied English terms.
- Avoid silently replacing one word with a different concept.

Example:

```text
{" desk ","cat","car"}
```

becomes conceptually:

```text
desk
cat
car
```

## 4.4 Validation

The script should detect:

- missing command-line argument;
- malformed set syntax;
- non-string entries;
- empty input;
- empty words.

The user should receive a clear error instead of a Python traceback whenever practical.

---

# 5. Phase 2 — Temporary Workspace

The script requires temporary storage because Internet images and processed transparent images must not remain as permanent project files.

## 5.1 Temporary Files

The temporary workspace should contain files such as:

```text
temporary/
    desk_original.*
    desk_transparent.png
    cat_original.*
    cat_transparent.png
    car_original.*
    car_transparent.png
```

The exact internal naming can vary, but filenames should be safe and traceable to the corresponding word.

## 5.2 Cleanup Requirement

At the end of the process:

- downloaded source images must be deleted;
- generated transparent intermediate images must be deleted;
- temporary directories should be removed when empty;
- the final output must remain.

The final output is:

```text
final_composition.png
```

---

# 6. Phase 3 — Internet Image Search

## 6.1 Search Requirement

For each input word, the script searches the Internet for an image representing that word.

The agreed implementation uses:

```text
duckduckgo_search
```

This avoids requiring a separate paid image-search API key.

## 6.2 Search Strategy

For each word:

```text
word
  ↓
image search
  ↓
candidate image URLs
  ↓
try candidates until a valid downloadable image is found
```

The script should not assume that the first returned result is always valid.

## 6.3 Candidate Validation

A candidate image should be rejected when:

- the URL cannot be downloaded;
- the response is not an image;
- the content is invalid or corrupted;
- Pillow cannot open the downloaded file;
- the image cannot be processed for background removal.

The script should continue trying another candidate when possible.

## 6.4 Search Failure

When no suitable image can be obtained for one word:

```text
Warning: no suitable image found for "word"
```

The script then continues with the remaining words instead of failing the entire composition.

This behavior is explicitly consistent with the current project README. fileciteturn0file0L20-L23

---

# 7. Phase 4 — Download Images

## 7.1 HTTP Download

The script uses:

```text
requests
```

for downloading image content.

## 7.2 Download Rules

Each image should be:

1. requested with a reasonable timeout;
2. checked for a successful response;
3. written to the temporary workspace;
4. opened and validated using Pillow.

## 7.3 Robustness

The script should handle:

- connection errors;
- timeout errors;
- invalid content;
- redirects;
- unsupported image formats;
- corrupted downloads.

A failure for one word must not unnecessarily stop processing of all other words.

---

# 8. Phase 5 — Background Removal

## 8.1 Technology

Use:

```text
rembg
```

for automatic background removal.

The required runtime dependency also includes:

```text
onnxruntime
```

The dependency list currently contains:

```text
duckduckgo_search
requests
pillow
rembg
onnxruntime
```

fileciteturn0file1L1-L5

## 8.2 Desired Result

Each source image should become a transparent PNG containing the main object.

Conceptually:

```text
Original:
[ object + original background ]

After processing:
[ object + transparent background ]
```

The transparency is important because the object will later be positioned over the common final canvas.

## 8.3 Model Initialization

The first execution may download the `rembg` model.

According to the current README, this first-time model download is expected, after which this background-removal step can work without Internet access. fileciteturn0file0L20-L22

## 8.4 Failure Handling

If background removal fails for a specific image:

- report a warning;
- discard that candidate;
- optionally try another search result;
- continue processing the other words whenever possible.

---

# 9. Phase 6 — Prepare Objects for Composition

After successful background removal, each object should be converted into a composition-ready transparent image.

## 9.1 Crop Unnecessary Transparent Space

Transparent margins should be minimized where practical so the object occupies its intended visual area.

This makes random placement and scaling more predictable.

## 9.2 Resize

Each object receives a suitable random scale.

The scaling logic must remain controlled so that:

- small objects do not become invisible;
- large objects do not dominate the entire canvas;
- all objects remain visually recognizable.

The exact scale range may be implemented as configurable constants.

## 9.3 Rotation

Apply a small random rotation to selected objects.

The purpose is to produce an organic collage-like arrangement rather than identical alignment.

Rotation should remain moderate to reduce accidental distortion.

---

# 10. Phase 7 — Final Canvas

## 10.1 Canvas Size

The final image must be:

```text
1024 × 1024 pixels
```

This is also the output size documented by the current README. fileciteturn0file0L16-L18

## 10.2 Background

The final canvas uses a:

```text
light warm beige / light beige background
```

The exact RGB/hex value should be centralized in one configuration constant so it can be adjusted easily later.

Example conceptual value:

```text
#F5E6D3
```

The exact color can be changed without redesigning the rest of the composition engine.

---

# 11. Phase 8 — Artistic Random Composition

## 11.1 Composition Objective

The output must look like one imaginative visual scene containing all selected objects.

It must not look like:

```text
[object 1] [object 2] [object 3]
```

in a rigid horizontal row.

Instead, the objects should be visually distributed with controlled randomness.

## 11.2 Random Placement

For each transparent object:

- choose a random position;
- ensure the object remains mostly inside the 1024×1024 canvas;
- allow controlled overlap;
- vary the scale;
- vary the rotation.

## 11.3 Artistic Randomness

Randomness should be constrained.

Uncontrolled randomness could result in:

- objects completely hidden behind another object;
- objects placed almost entirely outside the canvas;
- too much empty space;
- all objects occupying one tiny region.

Therefore the composition algorithm should use safe placement boundaries and reasonable overlap rules.

## 11.4 Overlap

A moderate amount of overlap is allowed and encouraged.

This helps the final image feel like a unified composition rather than separate stickers.

However:

- every successfully processed object should remain at least partially visible;
- overlap should not systematically hide smaller objects.

## 11.5 Ordering / Layering

Objects are pasted sequentially.

Layering may be randomized or otherwise varied so the final visual does not always place the same category of object in the foreground.

---

# 12. Phase 9 — Composition Quality Controls

The composition engine should include practical controls.

## 12.1 Boundary Safety

An object's calculated position should account for its final width and height.

The script should avoid placing the complete object far beyond the canvas.

## 12.2 Visibility

A successfully processed object should have a reasonable visible area.

A placement that completely hides an object should be avoided or retried.

## 12.3 Density

As the number of words increases, the composition should adapt.

For a small number of words:

```text
2–4 objects
```

more open space can be maintained.

For a larger number:

```text
5–10+ objects
```

the script should distribute the objects more carefully to avoid excessive overlap.

## 12.4 Random Seed

The implementation may use Python's `random` module.

A random seed can optionally be introduced later for reproducible compositions, but reproducibility is not required by the current specification.

---

# 13. Phase 10 — Save Final Image

The final composition must be saved as:

```text
final_composition.png
```

in the same directory from which the script is executed, matching the current README. fileciteturn0file0L16-L18

Recommended output properties:

```text
Format: PNG
Size:   1024 × 1024
Mode:   RGB/RGBA as appropriate for the final canvas
```

The final result is one image containing the successfully processed objects.

---

# 14. Phase 11 — Automatic Cleanup

Cleanup is mandatory.

After the final image has been successfully written:

```text
temporary downloaded images → delete
temporary transparent images → delete
temporary directory → remove when possible
```

The final file:

```text
final_composition.png
```

must not be deleted.

The current README explicitly states that both downloaded and background-removed temporary images are deleted automatically after completion. fileciteturn0file0L16-L18

---

# 15. Dependency Plan

The project uses the following Python packages:

```text
duckduckgo_search
requests
pillow
rembg
onnxruntime
```

These are already reflected in `requirements.txt`. fileciteturn0file1L1-L5

## Installation

The current documented installation command is:

```bash
pip install -r requirements.txt --break-system-packages
```

fileciteturn0file0L3-L5

The script should not require:

- OpenAI API;
- DALL·E API;
- Stability API;
- Google Images API key;
- any manually supplied AI image-generation key.

---

# 16. No External AI Image-Generation API

This project is specifically based on:

```text
Internet image search
        +
background removal
        +
programmatic composition
```

It is **not** based on generating a brand-new image from a text prompt through an external generative-image API.

Therefore no API key is part of the project requirements.

---

# 17. Error-Handling Strategy

## 17.1 Global Failures

The script should fail clearly for fundamental problems such as:

- no input argument;
- invalid input syntax;
- no valid words;
- inability to create the final canvas;
- inability to save the final output.

## 17.2 Per-Word Failures

A failure related to one word should be isolated when possible.

Example:

```text
desk  → success
cat   → success
car   → download failed
book  → success
```

The final composition should contain:

```text
desk + cat + book
```

and report a warning for `car`.

This matches the documented behavior that the script skips a word when it cannot find a suitable image and continues with the rest. fileciteturn0file0L20-L23

## 17.3 No Successful Objects

If every word fails and no object can be processed, the script should report the problem clearly rather than producing an empty misleading result.

---

# 18. User Experience in CMD

The command-line experience should be understandable.

A typical run can show progress similar to:

```text
Input words: desk, cat, car

[1/3] Searching image for: desk
      Downloaded image
      Removing background
      Object ready

[2/3] Searching image for: cat
      Downloaded image
      Removing background
      Object ready

[3/3] Searching image for: car
      Downloaded image
      Removing background
      Object ready

Composing final image...
Saved: final_composition.png

Cleaning temporary files...
Done.
```

Warnings should clearly identify the affected word.

---

# 19. Main Script Responsibilities

The main script should be organized into logical functions rather than one large block.

Recommended responsibilities:

```text
parse_input_words()
validate_words()
create_temp_workspace()
search_image_urls()
download_image()
remove_background()
prepare_object()
calculate_random_position()
compose_objects()
save_final_image()
cleanup_temp_files()
main()
```

The exact names may differ, but responsibilities should remain separated.

---

# 20. Recommended Processing Architecture

Conceptually:

```text
main()
 ├── parse input
 ├── validate input
 ├── create temporary workspace
 ├── for each word
 │    ├── search
 │    ├── download
 │    ├── validate image
 │    ├── remove background
 │    └── prepare object
 ├── create 1024×1024 canvas
 ├── compose prepared objects
 ├── save final_composition.png
 └── cleanup
```

This architecture makes the script easier to debug and modify.

---

# 21. File Structure

The intended simple project structure is:

```text
ProjectFolder/
│
├── surreal_compose.py
├── requirements.txt
├── README.md
├── MasterPlan.md
└── final_composition.png   ← generated after execution
```

Temporary image files should not remain in the project after a successful run.

---

# 22. Functional Requirements Checklist

The final implementation is considered aligned with this Master Plan when all of the following are true:

- [ ] Accepts input from CMD.
- [ ] Accepts the agreed format `{"word1","word2","word3"}`.
- [ ] Supports a variable number of words.
- [ ] Uses Internet image search.
- [ ] Does not require an image-generation API key.
- [ ] Downloads source images temporarily.
- [ ] Uses `rembg` to remove backgrounds.
- [ ] Produces transparent object images internally.
- [ ] Uses a 1024×1024 final canvas.
- [ ] Uses a light beige background.
- [ ] Uses artistic random placement.
- [ ] Uses controlled random scaling.
- [ ] Uses controlled random rotation.
- [ ] Allows controlled object overlap.
- [ ] Keeps processed objects reasonably visible.
- [ ] Creates one combined final image.
- [ ] Saves it as `final_composition.png`.
- [ ] Deletes temporary downloaded files.
- [ ] Deletes temporary background-removed files.
- [ ] Continues when an individual word cannot be processed.
- [ ] Reports meaningful warnings/errors in CMD.

---

# 23. Operational Notes

According to the current README:

### Installation

```bash
pip install -r requirements.txt --break-system-packages
```

### Example execution

```bash
python surreal_compose.py '{"desk","cat","car"}'
```

### Output

```text
final_composition.png
```

with:

```text
1024 × 1024
```

### Internet

Internet access is required for the image-search/download stage on every execution. fileciteturn0file0L20-L22

### First `rembg` Run

The first execution may download the required AI model for background removal. fileciteturn0file0L20-L21

---

# 24. Scope Boundaries

The following are **outside the current scope**:

1. Generating a new image through OpenAI/DALL·E or another external image-generation API.
2. Requiring a paid Google Images API.
3. Maintaining a permanent local library of downloaded source images.
4. Producing one final image per word.
5. Producing a rigid table/grid as the default layout.
6. Keeping temporary downloaded images after execution.
7. Requiring the user to manually remove image backgrounds.

---

# 25. Final Expected Result

Given:

```bash
python surreal_compose.py '{"desk","cat","car"}'
```

the expected workflow is:

```text
desk → Internet image → temporary download → background removal
cat  → Internet image → temporary download → background removal
car  → Internet image → temporary download → background removal

                 ↓

        1024×1024 beige canvas

                 ↓

     artistic random composition

                 ↓

       final_composition.png
```

The final image should visually read as a single imaginative composition containing the available objects, with the original image backgrounds removed and a unified light beige background.

All temporary source and intermediate files are then removed automatically.

---

# 26. Master Objective

The project should remain simple from the user's perspective:

```text
ONE CMD COMMAND
       ↓
SEARCH
       ↓
REMOVE BACKGROUNDS
       ↓
ARTISTIC COMPOSITION
       ↓
ONE FINAL PNG
```

The user supplies only the words.

The script handles image discovery, temporary storage, background removal, composition, output, and cleanup automatically.
