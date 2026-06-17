import React, { useState, useEffect, useRef } from 'react';
import { Dumbbell, Calendar, TrendingUp, Settings, RotateCcw, Camera, Download, Upload, Plus, Minus, Check, X, ChevronDown, ChevronRight, Youtube, Flame, Apple, Bike, Moon, Clock, Target, Edit3, Play, Square, Info, AlertCircle, Shield, Database, Copy, ClipboardPaste } from 'lucide-react';
import * as XLSX from 'xlsx';
import * as sync from './github-sync.js';

// ============================================================
// EXERCISE DETAILS
// ============================================================

const EXERCISE_DETAILS = {
  inc_press: {
    setup: ['Adjust seat so handles align with MID-CHEST', 'Bench angle 30° — not steeper', 'Feet flat, back firmly against pad', 'Pull shoulders DOWN and BACK before grabbing handles', 'Neutral or pronated grip'],
    execution: ['Unrack with arms extended, shoulders pinned back', 'Inhale, lower slowly toward upper chest (3 sec)', 'Stop when handles align with chest — DEEP stretch', 'Pause 1 sec at bottom without losing tension', 'Drive elbows forward and slightly together (1 sec up)', 'Stop just before lockout to keep tension on chest', 'Reset and repeat'],
    feel: ['Deep stretch across upper chest at bottom', 'Hard squeeze in upper chest as you press', 'Pressure in pecs, NOT front delts', 'Should NOT feel: neck tension, sharp shoulder pain'],
    mistakes: [
      { mistake: 'Bench >45° steepness', fix: '30° max — steeper trains front delts' },
      { mistake: 'Pressing toward face', fix: 'Press straight out from mid-chest' },
      { mistake: 'Bouncing weight off chest', fix: 'Pause 1 sec — preserves stretch' },
      { mistake: 'Elbows flared at 90°', fix: 'Tuck to ~45° — saves shoulders' },
    ],
    paragraph: 'The incline machine chest press at 30° is your foundational upper chest builder. The "press with elbows" cue keeps tension on the pecs rather than letting front delts steal the load. Machine version produces equivalent pec EMG to barbell (Schanke 2012, López-Vivancos 2023) with less injury risk.',
  },
  pec_deck: {
    setup: ['Seat height: handles at MID-CHEST when arms extended', 'Back firmly against pad, feet flat', 'Starting position lets arms stretch BEHIND shoulder line', 'Neutral grip (thumbs up) — safer for shoulders', 'Chest UP, shoulders back and down'],
    execution: ['Start with arms wide — feel the deep stretch', 'PAUSE 2 SECONDS in stretched position', 'Bring elbows together using chest, not arms', 'Squeeze hands together at front 1 sec', 'Slowly reverse 3 sec back to stretched position', 'Keep slight elbow bend throughout'],
    feel: ['Intense STRETCH across chest at open position', 'Hard SQUEEZE in mid-chest at closed position', 'Burning pump after 8-10 reps (metabolic stress)', 'Should NOT feel: shoulder pinching, bicep strain'],
    mistakes: [
      { mistake: 'Starting position not stretched enough', fix: 'Set machine so arms go behind shoulders' },
      { mistake: 'Rushing the eccentric', fix: 'Full 3-sec lowering — slow stretch is everything' },
      { mistake: 'Skipping the stretch pause', fix: '2-sec pause is non-negotiable for growth' },
    ],
    paragraph: 'Pec deck is high-ROI because it combines deep stretch under load (Maeo 2022 lengthened-position research) with peak contraction. 2-sec stretch pause is where most growth happens. Treat as a quality movement, not heavy.',
  },
  cable_fly_lh: {
    setup: ['Cables at LOWEST setting', 'Split stance, lean forward ~10°', 'Arms hanging at sides, palms forward', 'Slight fixed elbow bend (~15°)', 'Shoulder blades down and back'],
    execution: ['Sweep arms upward and inward in arcing motion', 'Imagine scooping water from below to above forehead', 'Bring hands together at chin/forehead — squeeze 1 sec', 'Slowly reverse 3 sec, keeping arms wide', 'Stop when DEEP stretch in upper chest', 'Pause 2 sec in stretched position', 'Keep elbow angle LOCKED throughout'],
    feel: ['Pulling stretch in UPPER chest at bottom', 'Strong upper chest contraction at top', 'Tension from sternum to shoulders'],
    mistakes: [
      { mistake: 'Bending elbows during rep', fix: 'Lock elbow angle, only shoulder moves' },
      { mistake: 'Not going low enough', fix: 'Let arms travel down/back for stretch' },
    ],
    paragraph: 'Low-to-high cable fly targets upper chest fibers via upward resistance vector. Cables maintain tension through full ROM unlike dumbbells. Bottom stretch position is the growth zone.',
  },
  cable_fly_hl: {
    setup: ['Cables at HIGHEST setting', 'Split stance, slight forward lean (~15°)', 'Arms raised above shoulder height', 'Palms forward, slight fixed elbow bend', 'Shoulder blades retracted'],
    execution: ['Sweep arms downward and inward (hug a barrel)', 'Meet hands at HIP level — not chest', 'Squeeze chest hard 1 sec at bottom', 'Slowly reverse 3 sec back up and wide', 'Deep stretch across chest at top', 'Keep elbow bend LOCKED'],
    feel: ['Pulling stretch across LOWER/MID chest at top', 'Strong contraction across chest at meet point', 'Tension from shoulders to sternum'],
    mistakes: [
      { mistake: 'Meeting at chest level', fix: 'Hands go to HIP — downward arc hits lower chest' },
      { mistake: 'Bending elbows', fix: 'Lock the angle, shoulder moves only' },
    ],
    paragraph: 'High-to-low fly targets lower/mid chest via downward resistance. Complements low-to-high fly for complete chest coverage. The arc (hug a barrel) creates different stimulus than presses.',
  },
  cable_lat_p: {
    setup: ['Cable at LOWEST setting', 'Stand sideways, cable BEHIND back at hip level', 'Cross-grip: far hand reaches across body', 'Stand tall, feet shoulder-width', 'BEFORE starting: press shoulder DOWN and BACK'],
    execution: ['Start: hand at opposite hip, arm across body', 'Lead with ELBOW — raise out to side and slightly back', 'Imagine pouring from a jug — elbow first', 'Raise until elbow at SHOULDER HEIGHT (not higher)', 'Pause 1 sec at top for peak contraction', 'Lower slowly 2 sec with tension', 'Reset shoulder before next rep'],
    feel: ['Burning on SIDE (lateral) of shoulder', 'Clean separation from traps', 'Pump after 12-15 reps'],
    mistakes: [
      { mistake: 'Shrugging — traps take over', fix: 'PRESS shoulder down first; drop weight if needed' },
      { mistake: 'Raising above shoulder height', fix: 'Stop at shoulder level' },
      { mistake: 'Cable in front of body', fix: 'Cable BEHIND for proper stretch position' },
    ],
    paragraph: 'Cable lateral raise from behind-back setup is gold standard for side delt hypertrophy. Wolf 2024: lengthened position drives most growth. Cable maintains tension at stretch position unlike dumbbells.',
  },
  cross_lat: {
    setup: ['Cable at LOWEST setting, single D-handle', 'Stand sideways with cable on far side', 'Grab handle by REACHING ACROSS body — starts at opposite hip', 'Free hand on hip/machine for balance', 'PRESS shoulder DOWN before starting'],
    execution: ['Start: hand at opposite hip, side delt in full stretch', 'Pause 1 sec in stretched start', 'Lead with ELBOW — raise up and out in Y shape', 'Motion is up AND across (sword from opposite hip)', 'Raise until elbow at shoulder height', 'Squeeze 1 sec at top', 'Slowly reverse 2 sec to cross-body stretched start'],
    feel: ['MASSIVE stretch in side delt at start', 'Clean isolation of shoulder side', 'Y-shape pull through full ROM'],
    mistakes: [
      { mistake: 'Starting at same-side hip', fix: 'Cross-body — opposite hip is the point' },
      { mistake: 'Raising hand not elbow', fix: 'Elbow leads, hand drags' },
    ],
    paragraph: 'Cross-body cable Y-raise is the most evidence-based side delt exercise (Nippard/Israetel). Cross-body start puts delt in MAX stretch — lengthened position research validates as highest stimulus.',
  },
  tri_oh: {
    setup: ['Cable HIGHEST, rope attached', 'Face AWAY, hands grip rope behind head', 'Step forward 2-3 steps so arms fully overhead', 'Stagger stance, elbows tucked to ears', 'Slight forward hip lean'],
    execution: ['Start with arms STRAIGHT overhead', 'Slowly bend elbows, hands drop behind neck (3 sec)', 'Lower until DEEP stretch in back of arms', 'PAUSE 2 SECONDS in stretched position', 'Press hands back overhead by extending elbows (1 sec)', 'Keep elbows pointing UP throughout', 'Flare rope apart slightly at top'],
    feel: ['Intense STRETCH along back of upper arm (long head)', 'Burning contraction at top lockout', 'Deep muscle sensation overhead'],
    mistakes: [
      { mistake: 'Elbows flaring wide', fix: 'Keep tight to ears' },
      { mistake: 'Cutting stretch short', fix: 'Let hands drop FULLY behind neck' },
      { mistake: 'Skipping stretch pause', fix: '2-sec pause is the whole exercise' },
    ],
    paragraph: 'Overhead tricep extension is the single best long-head builder. Long head crosses shoulder so only fully stretched when arm is overhead. Maeo 2022 + Wolf 2023: lengthened position = 1.5-2x growth.',
  },
  tri_oh_sa: {
    setup: ['Cable HIGHEST, single D-handle', 'Face AWAY, hand behind head (working side)', 'Step forward so arm fully overhead at stretch', 'Free hand on hip', 'Elbow points up, tucked to ear'],
    execution: ['Start arm straight overhead', 'Slowly bend elbow, hand drops behind head (3 sec)', 'PAUSE 2 SEC in deepest stretch', 'Press back to overhead lockout (1 sec)', 'Squeeze tricep at top'],
    feel: ['Deep stretch along back of one arm', 'Unilateral isolation reveals weak side'],
    mistakes: [
      { mistake: 'Elbow drifts forward', fix: 'Keep elbow point UP' },
    ],
    paragraph: 'Single-arm overhead extension allows greater ROM and identifies imbalances. Same lengthened-position benefits as rope version with unilateral focus.',
  },
  tri_pd: {
    setup: ['Cable HIGHEST, rope attached', 'Stand facing cable, slight forward lean ~10°', 'Neutral grip on rope (palms facing)', 'Elbows pinned to sides', 'Forearms parallel to floor at start'],
    execution: ['Press rope DOWN by extending elbow only', 'As you reach bottom, FLARE rope apart', 'Squeeze tricep HARD 1 sec at lockout', 'Slowly return up 2 sec', 'Stop at forearms parallel — don\'t let elbows drift'],
    feel: ['Strong contraction in OUTER (lateral) tricep head', 'Burning pump after 10-12 reps', 'Hands "pull apart" at bottom'],
    mistakes: [
      { mistake: 'Elbows drift forward', fix: 'Pin to ribs — drop weight if needed' },
      { mistake: 'Not flaring at bottom', fix: 'Rope flare = peak lateral head activation' },
    ],
    paragraph: 'Pressdown targets lateral head (horseshoe look). Pairs with overhead extension for complete tricep development. Rope flare at bottom is critical for peak lateral head contraction.',
  },
  tri_pd_sa: {
    setup: ['Cable HIGHEST, single D-handle', 'Stand sideways or facing, slight lean', 'Overhand grip, palm down', 'Elbow pinned to side', 'Free hand on hip'],
    execution: ['Press handle down by extending elbow only', 'Squeeze hard 1 sec at lockout', 'Slow return 2 sec to forearms parallel', 'Keep elbow stationary throughout'],
    feel: ['Lateral head contraction with different angle', 'Single-arm reveals strength imbalances'],
    mistakes: [
      { mistake: 'Wrist rotation cheating', fix: 'Keep wrist neutral, elbow extension only' },
    ],
    paragraph: 'Single-arm pressdown provides angle variation from rope version. Overhand grip biases lateral head differently than neutral rope grip.',
  },
  lat_pull_w: {
    setup: ['Wide neutral grip handle (palms facing each other)', 'Grip just outside shoulder-width', 'Sit, knee pad pinned to thighs', 'Reach up and grab handles, let weight stretch', 'Pull shoulders DOWN and BACK', 'Lean back ~15°, chest UP'],
    execution: ['Initiate by pulling ELBOWS DOWN toward back pockets', 'Pull until bar at top of chest', 'Squeeze lats HARD at bottom — elbows behind body', 'Pause 1 sec at bottom for peak contraction', 'Slowly let bar rise 2 sec to full extension', 'Allow lats to STRETCH at top'],
    feel: ['Strong pull down OUTSIDE of back (lats)', 'Deep stretch across lats at top', 'Hands feel like passive hooks'],
    mistakes: [
      { mistake: 'Pulling with hands/biceps', fix: 'Hands are hooks; pull from elbows' },
      { mistake: 'Leaning back too far', fix: '15° max — more becomes a row' },
      { mistake: 'Not letting lats stretch', fix: 'Let shoulders rise at top — full ROM' },
    ],
    paragraph: 'Wide neutral grip pulldown is primary lat-width builder. Neutral grip = full ROM with less shoulder stress. Wide grip emphasizes lat width specifically.',
  },
  cs_row: {
    setup: ['Chest pad supports sternum to upper abs', 'Foot platform: knees slightly bent', 'Chest pressed firmly into pad', 'Neutral grip on handles, arms extended', 'Brace core, take breath'],
    execution: ['Pull handles by driving ELBOWS back and slightly out', 'Squeeze shoulder blades HARD as handles approach sides', 'Pause 1 sec at peak contraction', 'Slowly extend arms 2 sec', 'Get FULL stretch in mid back at bottom', 'Chest stays glued to pad'],
    feel: ['Strong squeeze between shoulder blades at peak', 'Pulling sensation across mid back', 'Stretch across upper back at bottom'],
    mistakes: [
      { mistake: 'Chest lifts off pad', fix: 'Drop weight to keep chest pinned' },
      { mistake: 'Pulling with biceps', fix: 'Lead with elbows, hands passive' },
    ],
    paragraph: 'Chest-supported machine row eliminates lower back from the equation. You can train closer to failure with less stabilizer fatigue. Critical for mid-back thickness.',
  },
  sa_row: {
    setup: ['Cable at MID or LOW position (waist height)', 'Single D-handle attached', 'Sideways or facing, opposite foot forward', 'Reach forward with working arm to stretch lat', 'Free hand on hip or machine'],
    execution: ['Start in full reach — shoulder protracted, lat stretched', 'Pull handle to hip by driving ELBOW back past torso', 'Keep wrist neutral, hand passive', 'Squeeze lat HARD 1 sec at back position', 'Slowly reverse 2 sec to full reach/stretch', 'Allow shoulder to come forward — full lat stretch'],
    feel: ['Focused lat contraction on working side', 'Pulling/stretching from armpit to hip', 'Reveals side imbalances'],
    mistakes: [
      { mistake: 'Rotating torso', fix: 'Torso forward; only shoulder rotates' },
      { mistake: 'Stopping early', fix: 'Pull elbow PAST torso for full ROM' },
    ],
    paragraph: 'Single-arm cable row allows greater ROM than bilateral due to shoulder rotation. Identifies and corrects side-to-side imbalances. Use weaker side\'s performance to set both sides\' volume.',
  },
  cable_pull: {
    setup: ['Cable HIGHEST, straight bar or rope', 'Stand facing cable, 2-3 steps back', 'Hinge forward at hips ~30° — straight back', 'Palms-down grip, arms extended overhead', 'SLIGHT FIXED elbow bend — locked throughout'],
    execution: ['Pull bar down in arcing motion toward thighs', 'Squeeze armpits/lats — back, not arms', 'Drive bar to front of thighs', 'Squeeze lats 1 sec at bottom', 'Slowly reverse 3 sec back overhead', 'Allow lats to STRETCH fully at top', 'Elbow bend angle CONSTANT throughout'],
    feel: ['Pulling tension from armpits to hips', 'Clean lat isolation — no arms', 'Deep stretch when overhead'],
    mistakes: [
      { mistake: 'Elbows bending', fix: 'Lock the angle — only shoulder moves' },
      { mistake: 'Pulling with arms', fix: 'Squeeze armpits closed; lats initiate' },
    ],
    paragraph: 'Cable straight-arm pullover is the ONLY true lat isolation. Locked elbows mean biceps cannot help — every ounce of work goes through lats. Critical for building mind-muscle connection and adding stimulus after compound back work fatigued the biceps.',
  },
  rev_pec: {
    setup: ['Sit on pec deck FACING the pad', 'Seat height: handles at SHOULDER level', 'Neutral grip (thumbs up) — safer for shoulders', 'Chest pressed into pad, sit tall', 'Arms start in front, slightly bent at elbow'],
    execution: ['Pull handles back and out to sides, leading with ELBOWS', 'Hands are hooks — elbows drive', 'Pull until elbows align with shoulders', 'Squeeze rear delts HARD 1 sec at peak', 'Slowly reverse 2 sec back to start', 'Stop before handles touch — maintain tension'],
    feel: ['Burning in back of shoulders (rear delts)', 'Squeeze between shoulder blades at peak', 'Pump after 12-15 reps'],
    mistakes: [
      { mistake: 'Using arms to pull', fix: 'Soft grip, elbow drives' },
      { mistake: 'Pulling too far back', fix: 'Stop when elbows align with shoulders' },
    ],
    paragraph: 'Reverse pec deck is the most consistent rear delt builder. Machine path eliminates momentum cheating. Rear delts are critical for 3D shoulder development and posture.',
  },
  bayesian_curl: {
    setup: ['Cable LOWEST, single D-handle', 'Stand facing AWAY from cable, cable behind body', 'Step forward 1-2 steps so cable pulls arm BACK behind torso', 'Grab with one hand (single-arm movement)', 'Free hand on hip, stagger stance', 'Working arm: shoulder pulled back, elbow points down-and-back'],
    execution: ['Start with arm extended back behind body — feel stretch in bicep', 'PAUSE 2 SECONDS in stretched position', 'Curl handle to shoulder by bending elbow only', 'KEEP ELBOW STATIONARY behind body', 'Squeeze bicep hard 1 sec at top', 'Slowly reverse 3 sec back to stretched start', 'Complete all reps one arm, then switch'],
    feel: ['Profound STRETCH in bicep at start (unique sensation)', 'Hard contraction at top', 'Different/deeper burn than regular curls'],
    mistakes: [
      { mistake: 'Elbow drifting forward', fix: 'Keep elbow PINNED back; forearm only moves' },
      { mistake: 'Skipping stretch pause', fix: '2-sec pause is the whole exercise' },
      { mistake: 'Cable in front of body', fix: 'MUST be behind you for the stretch' },
    ],
    paragraph: 'Bayesian cable curl is the most effective bicep exercise in modern hypertrophy research. Maeo 2022 + Pedrosa 2022: lengthened-position curl produces ~2× the hypertrophy of standard curls. The 2-sec stretch pause is what makes this work.',
  },
  preacher: {
    setup: ['Adjust preacher bench: armpit snug into top of pad', 'Seat height: chest against bench, elbows fully supported', 'Cable handle or bar, supinated grip (palms up)', 'Pad locks elbows in position'],
    execution: ['Start with elbows extended (but NOT fully locked)', 'Curl handle up by bending elbow only', 'Squeeze bicep hard 1 sec at top', 'Slowly reverse 3 sec back down', 'STOP before full extension — keep tension', 'Maintain pad contact throughout'],
    feel: ['Strong contraction in SHORT HEAD of bicep (inner)', 'Burn develops faster than standing curls (constant tension)'],
    mistakes: [
      { mistake: 'Fully extending at bottom', fix: 'Stop short — protects elbows, keeps tension' },
      { mistake: 'Shoulder swinging weight', fix: 'Elbows locked in pad; only forearm moves' },
    ],
    paragraph: 'Cable preacher curl isolates the bicep short head (inner peak). Preacher bench eliminates shoulder/back cheating. Cable adds constant tension that dumbbell version lacks at the top.',
  },
  hammer_c: {
    setup: ['Cable LOWEST, rope attached', 'Stand facing cable, 1 step back', 'Neutral grip (palms facing each other)', 'Elbows pinned to sides — won\'t move', 'Stand tall, slight forward lean'],
    execution: ['Curl rope to shoulders by bending elbows only', 'Maintain neutral grip throughout', 'At top, TWIST rope OUTWARD (pull hands apart)', 'Hold 1 sec at top', 'Slowly reverse 2 sec to full extension', 'Slight stretch at bottom, elbow position locked'],
    feel: ['Strong contraction in BRACHIALIS (outside of upper arm)', 'Secondary bicep/forearm engagement', 'Fuller upper arm sensation'],
    mistakes: [
      { mistake: 'Not twisting at top', fix: 'Rope flare = peak brachialis activation' },
      { mistake: 'Using body swing', fix: 'Lean for balance only, no momentum' },
    ],
    paragraph: 'Cable hammer curl targets the brachialis — the muscle under the bicep that pushes it up visually. Neutral grip biases brachialis over bicep. Rope twist at top is the secret cue.',
  },
  flat_press: {
    setup: ['Seat height: handles align with mid-chest', 'Feet flat, back firmly against pad', 'Pull shoulders back and DOWN before grabbing', 'Neutral or pronated grip', 'Breath, brace core'],
    execution: ['Unrack with arms extended, shoulders pinned', 'Lower slowly 3 sec to mid-chest level', 'Feel STRETCH across pecs at bottom', 'Pause 1 sec at bottom', 'Press by driving elbows forward and together', 'Stop just before lockout'],
    feel: ['Stretch across mid-chest at bottom', 'Squeeze across entire chest pressing up', 'Pressure in pecs, NOT front delts'],
    mistakes: [
      { mistake: 'Handles too high (toward face)', fix: 'Align with mid-chest' },
      { mistake: 'Bouncing off chest', fix: 'Pause 1 sec at bottom' },
    ],
    paragraph: 'Flat machine press hits mid-chest fibers (sternal portion) more directly than incline. Used as second weekly chest hit, provides different stimulus angle for complete chest development.',
  },
  shoulder_p: {
    setup: ['Seat height: handles align with SHOULDERS', 'Back firmly against pad, slight backward lean (~5°)', 'Feet flat for stability', 'Neutral or pronated grip', 'Pull shoulders DOWN, chest UP'],
    execution: ['Press by driving ELBOWS up (not hands forward)', 'Stop just before lockout — tension on delts', 'Slowly lower 3 sec to shoulder/ear level', 'Pause 1 sec at bottom', 'Drive elbows up again — explosive but controlled'],
    feel: ['Burn across front and side of shoulders', 'Secondary tricep engagement', 'Should NOT feel: dominant traps, lower back strain'],
    mistakes: [
      { mistake: 'Pressing hands forward', fix: 'Drive through elbows, hands follow' },
      { mistake: 'Locking out at top', fix: 'Stop short to keep delts loaded' },
    ],
    paragraph: 'Machine shoulder press is the primary front delt builder with side delt involvement. Seated position with back support eliminates lower back. "Drive elbows up" engages more of the delt as a whole.',
  },
  face_pull: {
    setup: ['Cable slightly ABOVE head height', 'Rope attached, palms facing each other (thumbs back)', 'Step back 2-3 feet, cable loaded at start', 'Stand tall, hip-width feet, slight forward lean'],
    execution: ['Pull rope TOWARD FOREHEAD by driving elbows UP and OUT', 'At end of pull, EXTERNALLY ROTATE — hands above elbows', 'Squeeze rear delts and mid traps 1 sec', 'Slowly extend 2 sec back to start', 'Maintain HIGH elbows throughout'],
    feel: ['Burn in REAR DELTS and across mid back', 'Scapular squeeze between shoulder blades'],
    mistakes: [
      { mistake: 'Low elbows (pull to chest)', fix: 'Elbows STAY UP; pull to forehead' },
      { mistake: 'No external rotation', fix: 'Hands above elbows at end — critical cue' },
    ],
    paragraph: 'Cable face pull hits rear delts, mid traps, and rotator cuff externals all at once. Best single "posture" and shoulder health exercise. High elbows + external rotation are the two critical cues.',
  },
  leg_press: {
    setup: ['Sit with lower back FLAT against pad', 'Feet shoulder-width on platform, mid-height', 'Toes pointed slightly outward (~15°)', 'Unrack, lower carriage to start', 'Knees track in line with toes'],
    execution: ['Lower platform slowly 3 sec by bending hips and knees', 'Knees TOWARD chest (full ROM, not partial)', 'STOP when lower back about to round', 'Drive platform up through whole foot', 'Stop just before knee lockout'],
    feel: ['Strong quad burn (front of thighs)', 'Secondary glute engagement', 'Should NOT feel: lower back rounding'],
    mistakes: [
      { mistake: 'Partial reps', fix: 'Full ROM — knees toward chest' },
      { mistake: 'Lower back lifts off pad', fix: 'Stop short of that point' },
    ],
    paragraph: 'Leg press is your foundational lower body compound. Safer than squats for lumbar spine while building quads primarily. Full ROM is non-negotiable for hypertrophy.',
  },
  seated_curl: {
    setup: ['Adjust seat: back of knees against front edge', 'Leg pad on BACK of ankles', 'Lap belt secured — hips anchored', 'Sit tall, chest up, hands on side handles', 'Start with legs nearly straight'],
    execution: ['Curl heels DOWN and toward glutes by bending knees', 'POINT TOES UP throughout — biases hams over calves', 'Pull until knees bent ~90°', 'Squeeze hamstrings 1 sec at bottom', 'Slowly extend 3 sec back to start', 'Feel stretch in hams', 'Don\'t let weight stack touch'],
    feel: ['Burn in HAMSTRINGS', 'Stretch across hams at top', 'Should NOT feel: calf cramping or lower back strain'],
    mistakes: [
      { mistake: 'Toes pointed down', fix: 'Point UP to bias hams over calves' },
      { mistake: 'Hips lifting off seat', fix: 'Use lap belt or drop weight' },
    ],
    paragraph: 'Seated leg curl is now considered the BEST hamstring exercise, surpassing lying curls (Maeo 2021). Seated position flexes hip, lengthening the hamstring. Lengthened position = more hypertrophy. Toes UP is critical to prevent calf takeover.',
  },
  leg_ext: {
    setup: ['Adjust seat: knees align with machine pivot', 'Shin pad on LOWER shins (above ankle)', 'Back firmly against seat, hands on handles', 'Sit tall, core engaged'],
    execution: ['Extend legs by squeezing quads', 'Point toes slightly UP at top for max contraction', 'Pause 1 sec at top', 'Slowly lower 2 sec to ~90° knee bend', 'Don\'t let weight stack touch — constant tension'],
    feel: ['Sharp quad burn, especially inner thigh (VMO)', 'Peak contraction at top'],
    mistakes: [
      { mistake: 'Swinging weight up', fix: 'Slow controlled, quads only' },
      { mistake: 'Pad on top of foot', fix: 'On shin above ankle' },
    ],
    paragraph: 'Leg extension is direct quad isolator. Toes up cue maximizes VMO (teardrop above knee). 2 sets sufficient given moderate leg goal.',
  },
  calf: {
    setup: ['Balls of feet on platform', 'Heels hanging off — need drop for stretch', 'Shoulders/knees under pad (standing or seated)', 'Stand tall, core engaged'],
    execution: ['Lower heels SLOWLY 2 sec below platform', 'PAUSE 2 SECONDS at deep stretch bottom', 'Push up onto balls of feet by squeezing calves', 'Rise as high as possible — full plantarflexion', 'Squeeze calves 1 sec at top', 'Slow deliberate reps'],
    feel: ['Deep stretch in calves at bottom', 'Hard squeeze at top', 'Burning pump that lingers'],
    mistakes: [
      { mistake: 'Bouncing reps', fix: 'Controlled tempo + pauses' },
      { mistake: 'No heel drop', fix: 'Stretch IS the exercise' },
    ],
    paragraph: 'Calves are stubborn because most people train them poorly. Solution: full ROM + slow tempo + 2-sec stretch pause. Calves are 60-80% slow-twitch — they respond to time under tension at long lengths.',
  },
  knee_raise: {
    setup: ['Hang from pull-up bar, arms straight', 'Use straps if grip limits', 'Engage core, shoulders down and back', 'Legs together, slight knee bend'],
    execution: ['Curl PELVIS up by squeezing abs (not just lifting legs)', 'Bring knees toward chest', 'Pause 1 sec at top with knees high', 'Slowly lower 2 sec — control descent', 'No body swing — full stop between reps'],
    feel: ['Strong contraction across entire abdominal wall', 'Pelvic tilt sensation — the key'],
    mistakes: [
      { mistake: 'Swinging legs without pelvic curl', fix: 'It\'s a CRUNCH hanging, not leg lift' },
    ],
    paragraph: 'Hanging knee raises are the most effective ab exercise. Key: this is a CRUNCH performed hanging — pelvis curls up toward ribcage, not legs swinging. Curl the pelvis to fight anterior pelvic tilt.',
  },
  hack_squat: {
    setup: ['Back and shoulders flat against the pad', 'Feet shoulder-width, mid-platform, toes slightly out', 'Brace core, unlock the safeties', 'Knees track in line with toes'],
    execution: ['Lower slowly 3 sec, knees travelling over toes', 'Go as deep as control allows (thighs past parallel)', 'No bouncing — pause briefly at the bottom', 'Drive up through the whole foot', 'Stop just before knee lockout to keep tension'],
    feel: ['Strong quad burn (front of thighs)', 'Deep stretch at the bottom', 'Should NOT feel: lower-back rounding or knee pinching'],
    mistakes: [
      { mistake: 'Quarter reps', fix: 'Full depth — the stretch builds the quad' },
      { mistake: 'Heels lifting', fix: 'Push through mid-foot; lower the load' },
    ],
    paragraph: 'The hack squat machine is a back-supported, quad-biased compound — a safe second quad angle alongside the leg press. Deep ROM under load drives quad growth without loading the spine.',
  },
  lying_curl: {
    setup: ['Lie face-down, knees just off the pad edge', 'Ankle roller on the BACK of the ankles', 'Hips pressed into the bench', 'Grip the handles, point toes down (or neutral)'],
    execution: ['Curl heels toward glutes by bending the knees', 'Squeeze hamstrings hard 1 sec at the top', 'Slowly lower 3 sec to a stretch', 'Keep hips DOWN — no lifting off the pad', 'Stop short of the stack touching'],
    feel: ['Hamstring contraction at the top', 'Stretch down the back of the thigh', 'Should NOT feel: lower-back arching or cramping'],
    mistakes: [
      { mistake: 'Hips lifting off the pad', fix: 'Press hips down; drop weight if needed' },
      { mistake: 'Rushing the lowering', fix: 'Slow 3-sec eccentric is the growth driver' },
    ],
    paragraph: 'Lying leg curl trains the hamstrings at a shorter hip angle than the seated curl — a complementary second angle. Slow eccentrics and a hard top squeeze are what matter for the moderate leg goal.',
  },
  rev_grip_pull: {
    setup: ['Narrow UNDERHAND grip (palms facing you), about shoulder-width', 'Use your versa-grip straps so forearms/grip never limit the lats', 'Sit, knee pad snug, lean back ~10–15°', 'Pull shoulders DOWN and back; chest up'],
    execution: ['Initiate by driving the ELBOWS DOWN and into your ribs', 'Think "elbows to back pockets" — hands are just hooks (straps help)', 'Pull the bar to upper chest, squeeze lats 1 sec', 'Slowly let it rise 2–3 sec to a full lat stretch', 'Keep the chest up — do not let the torso collapse forward'],
    feel: ['Strong contraction LOW on the lats (thickness)', 'Stretch across the lats at the top', 'Should NOT be a biceps or forearm exercise — straps + elbow drive keep it on the back'],
    mistakes: [
      { mistake: 'Curling with the biceps', fix: 'Straps on, lead with elbows — hands passive' },
      { mistake: 'Leaning back to heave it', fix: 'Stay tall, ~15° max, control the weight' },
    ],
    paragraph: 'The reverse-grip (underhand) pulldown replaces the straight-arm pullover you disliked (it kept turning into a triceps exercise). The underhand grip + elbow-drive + versa-grip straps put the work on the lower lats — back, not arms. Done on your Panatta converging plate-loaded pulldown.',
  },
};

// ============================================================
// PROGRAM DATA — V4 100/100 PROGRAM
// ============================================================

const PROGRAM = {
  push: {
    label: 'PUSH', name: 'PUSH', focus: 'Chest · Shoulders · Triceps',
    mobility: ['Band pull-aparts × 15', 'Wall slides × 10', 'Shoulder dislocates × 10'],
    exercises: [
      { id: 'pec_deck', name: 'Pec Deck', garmin: 'Pec Deck', target: 'Chest (stretch)', sets: 3, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-1', rest: 75, cue: '2-sec pause at full stretch. Lead chest, not arms. Last 1-2 sets to 0-1 RIR.', weightMode: 'total', startWeight: 45 },
      { id: 'cable_fly_lh', name: 'Cable Fly (low-to-high)', garmin: 'Cable Fly Low To High', target: 'Upper chest (stretch)', sets: 2, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-0', rest: 75, cue: 'Fixed elbow angle; meet hands at chin level.', weightMode: 'total', startWeight: 10 },
      { id: 'inc_press', name: 'Incline Machine Chest Press (30°)', garmin: 'Incline Machine Chest Press', target: 'Upper chest', sets: 3, repMin: 6, repMax: 12, rirMin: 1, rirMax: 2, tempo: '3-1-1-0', rest: 150, cue: 'Drive elbows forward, not hands. Feel chest, not triceps.', weightMode: 'total', startWeight: 38 },
      { id: 'shoulder_p', name: 'Machine Shoulder Press', garmin: 'Machine Shoulder Press', target: 'Front + side delts', sets: 3, repMin: 6, repMax: 12, rirMin: 1, rirMax: 2, tempo: '3-1-1-0', rest: 150, cue: 'Drive elbows UP, not hands forward.', weightMode: 'total', startWeight: 25 },
      { id: 'cross_lat', name: 'Cross-Cable Lateral Raise', garmin: 'Cross-Cable Lateral Raise', target: 'Side delts (max stretch)', sets: 4, repMin: 12, repMax: 15, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 75, cue: 'Hand at OPPOSITE hip. Lead with elbow.', weightMode: 'perSide', startWeight: 5 },
      { id: 'tri_pd', name: 'Triceps Pushdown (rope)', garmin: 'Triceps Pushdown', target: 'Triceps lateral head', sets: 3, repMin: 10, repMax: 15, rirMin: 1, rirMax: 2, tempo: '2-1-1-1', rest: 75, cue: 'Pin elbows. Flare rope apart at bottom.', weightMode: 'total', startWeight: 20 },
      { id: 'tri_oh', name: 'Overhead Cable Triceps Extension', garmin: 'Overhead Cable Triceps Extension', target: 'Triceps long head', sets: 2, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-0', rest: 75, cue: '2-sec stretch overhead. Elbows tucked to ears.', weightMode: 'total', startWeight: 30 },
    ],
  },
  pull: {
    label: 'PULL', name: 'PULL', focus: 'Back · Rear delts · Biceps',
    mobility: ['Scapular pulls × 10', 'Band pull-aparts × 15', 'Cat-cow × 10'],
    exercises: [
      { id: 'lat_pull_w', name: 'Wide Lat Pulldown', garmin: 'Wide Lat Pulldown', target: 'Lats — width', sets: 4, repMin: 8, repMax: 12, rirMin: 1, rirMax: 2, tempo: '2-1-1-0', rest: 150, cue: 'Pull elbows DOWN toward back pockets.', weightMode: 'total', startWeight: 45 },
      { id: 'cs_row', name: 'Chest-Supported Row', garmin: 'Chest-Supported Row', target: 'Mid back, lats', sets: 4, repMin: 8, repMax: 12, rirMin: 1, rirMax: 2, tempo: '2-1-1-1', rest: 150, cue: 'Drive elbows back, squeeze 1 sec.', weightMode: 'total', startWeight: 35 },
      { id: 'sa_row', name: 'Single-Arm Cable Row', garmin: 'Single-Arm Cable Row', target: 'Back unilateral', sets: 3, repMin: 10, repMax: 12, rirMin: 1, rirMax: 2, tempo: '2-1-1-1', rest: 90, cue: 'Reach forward to stretch. Pull elbow past torso.', weightMode: 'perSide', startWeight: 20 },
      { id: 'rev_grip_pull', name: 'Reverse-Grip Lat Pulldown', garmin: 'Reverse-Grip Lat Pulldown', target: 'Lats — lower / thickness', sets: 3, repMin: 10, repMax: 12, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 90, cue: 'Underhand narrow grip + versa straps. Drive elbows to ribs — lats, not arms.', weightMode: 'total', startWeight: 38 },
      { id: 'rev_pec', name: 'Reverse Pec Deck', garmin: 'Reverse Pec Deck', target: 'Rear delts', sets: 3, repMin: 12, repMax: 15, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 60, cue: 'Lead with elbows. Soft grip.', weightMode: 'total', startWeight: 37 },
      { id: 'bayesian_curl', name: 'Bayesian Cable Curl', garmin: 'Bayesian Cable Curl', target: 'Biceps (stretch)', sets: 4, repMin: 10, repMax: 15, rirMin: 1, rirMax: 2, tempo: '3-2-1-1', rest: 75, cue: 'Cable BEHIND body. 2-sec stretch pause. Per side.', weightMode: 'perSide', startWeight: 7 },
      { id: 'preacher', name: 'Cable Preacher Curl', garmin: 'Preacher Curl', target: 'Biceps short head', sets: 3, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-1-1-1', rest: 75, cue: "Don't lock at bottom. Slow eccentric.", weightMode: 'total', startWeight: 22 },
    ],
  },
  lower: {
    label: 'LOWER', name: 'LOWER (moderate)', focus: 'Quads · Hams · Calves · Abs',
    mobility: ['Bodyweight squats × 15', 'Glute bridges × 15', 'Hip circles × 10/side', 'Leg swings × 10/dir'],
    exercises: [
      { id: 'hack_squat', name: 'Hack Squat Machine', garmin: 'Hack Squat Machine', target: 'Quads', sets: 3, repMin: 8, repMax: 12, rirMin: 1, rirMax: 2, tempo: '3-1-1-0', rest: 150, cue: 'Full depth, knees over toes. Drive through mid-foot.', weightMode: 'total', startWeight: 50 },
      { id: 'leg_press', name: 'Leg Press', garmin: 'Leg Press', target: 'Quads, glutes', sets: 3, repMin: 10, repMax: 15, rirMin: 1, rirMax: 2, tempo: '3-1-1-0', rest: 150, cue: 'Full ROM, knees toward chest.', weightMode: 'total', startWeight: 70 },
      { id: 'leg_ext', name: 'Leg Extension', garmin: 'Leg Extension', target: 'Quads (isolation)', sets: 2, repMin: 12, repMax: 15, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 75, cue: '1-sec squeeze at top, toes up.', weightMode: 'total', startWeight: 30 },
      { id: 'seated_curl', name: 'Seated Leg Curl', garmin: 'Seated Leg Curl', target: 'Hamstrings', sets: 3, repMin: 10, repMax: 15, rirMin: 1, rirMax: 2, tempo: '3-1-1-1', rest: 75, cue: 'Toes UP to bias hams. Slow eccentric.', weightMode: 'total', startWeight: 37 },
      { id: 'calf', name: 'Calf Raise', garmin: 'Calf Raise', target: 'Calves', sets: 4, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '2-2-1-1', rest: 75, cue: '2-sec stretch at bottom. Full rise.', weightMode: 'total', startWeight: 80 },
      { id: 'knee_raise', name: 'Hanging Knee Raise', garmin: 'Hanging Knee Raise', target: 'Abs (anti-tilt)', sets: 3, repMin: 12, repMax: 20, rirMin: 0, rirMax: 2, tempo: '2-1-1-1', rest: 75, cue: 'Curl pelvis UP, not just legs.', weightMode: 'total', startWeight: 0 },
    ],
  },
  upper: {
    label: 'UPPER', name: 'UPPER (2nd hit)', focus: 'Chest · Back · Delts · Arms — fresh angles',
    mobility: ['Band pull-aparts × 15', 'Wall slides × 10', 'Cat-cow × 10'],
    exercises: [
      { id: 'pec_deck', name: 'Pec Deck', garmin: 'Pec Deck', target: 'Chest (stretch)', sets: 3, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-1', rest: 75, cue: '2-sec pause at full stretch. Feel the chest. Last 1-2 sets to 0-1 RIR.', weightMode: 'total', startWeight: 45 },
      { id: 'cable_fly_hl', name: 'Cable Fly (high-to-low)', garmin: 'Cable Fly High To Low', target: 'Lower/mid chest (stretch)', sets: 3, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-0', rest: 75, cue: 'Hug a barrel down to hips. Fixed elbow — chest, zero triceps.', weightMode: 'total', startWeight: 10 },
      { id: 'sa_row', name: 'Single-Arm Cable Row', garmin: 'Single-Arm Cable Row', target: 'Back unilateral', sets: 4, repMin: 10, repMax: 12, rirMin: 1, rirMax: 2, tempo: '2-1-1-1', rest: 90, cue: 'Reach forward to stretch. Pull elbow past torso.', weightMode: 'perSide', startWeight: 20 },
      { id: 'lat_pull_w', name: 'Wide Lat Pulldown', garmin: 'Wide Lat Pulldown', target: 'Lats — width', sets: 2, repMin: 8, repMax: 12, rirMin: 1, rirMax: 2, tempo: '2-1-1-0', rest: 120, cue: 'Pull elbows DOWN toward back pockets.', weightMode: 'total', startWeight: 45 },
      { id: 'tri_oh_sa', name: 'Single-Arm Overhead Triceps', garmin: 'Single-Arm Overhead Triceps', target: 'Triceps long head', sets: 3, repMin: 10, repMax: 15, rirMin: 0, rirMax: 1, tempo: '3-2-1-0', rest: 75, cue: '2-sec stretch overhead.', weightMode: 'perSide', startWeight: 10 },
      { id: 'hammer_c', name: 'Cable Hammer Curl (rope)', garmin: 'Cable Hammer Curl', target: 'Brachialis', sets: 3, repMin: 10, repMax: 15, rirMin: 1, rirMax: 2, tempo: '2-1-1-1', rest: 75, cue: 'Twist rope outward at top.', weightMode: 'total', startWeight: 12 },
      { id: 'cable_lat_p', name: 'Cable Lateral Raise', garmin: 'Cable Lateral Raise', target: 'Side delts', sets: 4, repMin: 12, repMax: 15, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 75, cue: 'Cable BEHIND body. Press shoulder down first.', weightMode: 'perSide', startWeight: 5 },
      { id: 'face_pull', name: 'Cable Face Pull', garmin: 'Face Pull', target: 'Rear delts, mid traps', sets: 3, repMin: 12, repMax: 20, rirMin: 0, rirMax: 1, tempo: '2-1-1-1', rest: 75, cue: 'High elbows. Pull to forehead. Externally rotate.', weightMode: 'total', startWeight: 10 },
    ],
  },
};

const DAYS_ORDER = ['push', 'pull', 'lower', 'upper'];

// ============================================================
// UTILITIES
// ============================================================

const todayKey = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

// A workout is keyed by (date + day) so two different days on the same date
// never share one record. This is the fix for the "switching days corrupts/
// orphans my logs" bug.
const sessionKey = (date, day) => `${date}__${day}`;

// Migrate legacy date-only keyed history (v4) to the (date+day) session model,
// preserving all old sessions. Old finished workouts get finishedAt = endTime.
const migrateHistory = (h) => {
  if (!h) return {};
  const out = {};
  for (const [k, rec] of Object.entries(h)) {
    if (/^\d{4}-\d{2}-\d{2}$/.test(k)) {
      const day = rec.day || 'push';
      out[sessionKey(k, day)] = { ...rec, date: k, day, finishedAt: rec.finishedAt || rec.endTime || null };
    } else {
      out[k] = rec.date ? rec : { ...rec, date: k.split('__')[0] };
    }
  }
  return out;
};

const formatTime = (seconds) => {
  if (seconds < 0) seconds = 0;
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = Math.floor(seconds % 60);
  if (h > 0) return `${h}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
  return `${m}:${String(s).padStart(2, '0')}`;
};

const formatClockTime = (iso) => {
  const d = new Date(iso);
  return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}:${String(d.getSeconds()).padStart(2, '0')}`;
};

const kgToLb = (kg) => Math.round(kg * 2.20462 * 10) / 10;
const lbToKg = (lb) => Math.round(lb / 2.20462 * 10) / 10;

const compressImage = (file) => new Promise((resolve) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const maxDim = 400;
      let w = img.width, h = img.height;
      if (w > h) { if (w > maxDim) { h = h * (maxDim / w); w = maxDim; } }
      else { if (h > maxDim) { w = w * (maxDim / h); h = maxDim; } }
      canvas.width = w; canvas.height = h;
      canvas.getContext('2d').drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', 0.7));
    };
    img.src = e.target.result;
  };
  reader.readAsDataURL(file);
});

const suggestWeight = (lastSession, repMin, repMax, rirTarget) => {
  if (!lastSession || !lastSession.sets || lastSession.sets.length === 0) return null;
  const sets = lastSession.sets;
  const allHitTop = sets.every(s => s.reps >= repMax && s.rir >= rirTarget);
  const lastWeight = sets[sets.length - 1].weight;
  const unit = sets[sets.length - 1].unit;
  if (allHitTop) return { weight: lastWeight + (unit === 'kg' ? 2.5 : 5), unit, reason: 'Top of range hit. Add weight.' };
  return { weight: lastWeight, unit, reason: 'Same weight. Push for more reps.' };
};

// ============================================================
// MAIN APP
// ============================================================

export default function HypertrophyApp() {
  const [view, setView] = useState('today');
  const [selectedDay, setSelectedDay] = useState('push');
  const [settings, setSettings] = useState({ unit: 'kg', bodyweight: 73, mesoWeek: 1, age: 30 });
  const [history, setHistory] = useState({});
  const [photos, setPhotos] = useState({});
  const [videos, setVideos] = useState({});
  const [bodyweightLog, setBodyweightLog] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [lastBackup, setLastBackup] = useState(null);

  const [restEndAt, setRestEndAt] = useState(null);
  const [restTotalSec, setRestTotalSec] = useState(0);
  const restNotifiedRef = useRef(false);

  const [nowTick, setNowTick] = useState(0);
  const [toast, setToast] = useState(null);
  const [zoomPhoto, setZoomPhoto] = useState(null);

  // GitHub two-way sync
  const [syncStatus, setSyncStatus] = useState(sync.isConfigured() ? 'idle' : 'off'); // off|idle|syncing|synced|offline|error
  const [lastSyncAt, setLastSyncAt] = useState(null);
  const shaRef = useRef({ history: null, photos: null });
  const syncReadyRef = useRef(false);
  const pushTimer = useRef(null);
  const photoTimer = useRef(null);

  // Session timing now lives INSIDE each (date+day) record, so each workout is
  // its own session. These derive the active session for the selected day.
  const currentKey = sessionKey(todayKey(), selectedDay);
  const currentRec = history[currentKey];
  const sessionStart = currentRec?.startTime || null;
  const sessionEnd = currentRec?.finishedAt || null;

  // Load on mount
  useEffect(() => {
    (async () => {
      const loadKey = async (key) => {
        try {
          const r = await window.storage.get(key);
          return r ? JSON.parse(r.value) : null;
        } catch (e) { return null; }
      };
      const s = await loadKey('hyp_settings'); if (s) setSettings(prev => ({ ...prev, ...s }));
      const h = await loadKey('hyp_history'); if (h) setHistory(migrateHistory(h));
      const p = await loadKey('hyp_photos'); if (p) setPhotos(p);
      const v = await loadKey('hyp_videos'); if (v) setVideos(v);
      const bw = await loadKey('hyp_bodyweight'); if (bw) setBodyweightLog(bw);
      const rest = await loadKey('hyp_rest'); if (rest && rest.restEndAt && rest.restEndAt > Date.now()) { setRestEndAt(rest.restEndAt); setRestTotalSec(rest.restTotalSec); }
      const lb = await loadKey('hyp_last_backup'); if (lb) setLastBackup(lb);
      setLoaded(true);
    })();
  }, []);

  useEffect(() => { if (loaded) window.storage.set('hyp_settings', JSON.stringify(settings)).catch(() => {}); }, [settings, loaded]);
  useEffect(() => { if (loaded) window.storage.set('hyp_history', JSON.stringify(history)).catch(() => {}); }, [history, loaded]);
  useEffect(() => { if (loaded) window.storage.set('hyp_photos', JSON.stringify(photos)).catch(() => {}); }, [photos, loaded]);
  useEffect(() => { if (loaded) window.storage.set('hyp_videos', JSON.stringify(videos)).catch(() => {}); }, [videos, loaded]);
  useEffect(() => { if (loaded) window.storage.set('hyp_bodyweight', JSON.stringify(bodyweightLog)).catch(() => {}); }, [bodyweightLog, loaded]);
  useEffect(() => { if (loaded && restEndAt) window.storage.set('hyp_rest', JSON.stringify({ restEndAt, restTotalSec })).catch(() => {}); }, [restEndAt, restTotalSec, loaded]);

  // --- GitHub sync: pull + merge on load ---
  useEffect(() => {
    if (!loaded || !sync.isConfigured()) return;
    (async () => {
      setSyncStatus('syncing');
      try {
        const local = { settings, history, photos, videos, bodyweightLog, updatedAt: localStorage.getItem('hyp_updatedAt') || '', hasLocal: Object.keys(history).length > 0 };
        const m = await sync.pullMerge(local);
        shaRef.current = m.shas;
        setSettings(prev => ({ ...prev, ...m.settings }));
        setHistory(m.history); setBodyweightLog(m.bodyweightLog); setVideos(m.videos); setPhotos(m.photos);
        if (m.localAhead) shaRef.current.history = await sync.pushHistory({ settings: m.settings, history: m.history, bodyweightLog: m.bodyweightLog, videos: m.videos }, m.shas.history);
        setSyncStatus('synced'); setLastSyncAt(new Date().toISOString());
      } catch (e) { setSyncStatus(navigator.onLine ? 'error' : 'offline'); }
      setTimeout(() => { syncReadyRef.current = true; }, 0);
    })();
  }, [loaded]);

  // --- push history.json on change (debounced) ---
  useEffect(() => {
    if (!syncReadyRef.current || !sync.isConfigured()) return;
    clearTimeout(pushTimer.current);
    pushTimer.current = setTimeout(async () => {
      setSyncStatus('syncing');
      try {
        shaRef.current.history = await sync.pushHistory({ settings, history, bodyweightLog, videos }, shaRef.current.history);
        localStorage.setItem('hyp_updatedAt', new Date().toISOString());
        setSyncStatus('synced'); setLastSyncAt(new Date().toISOString());
      } catch (e) { setSyncStatus(navigator.onLine ? 'error' : 'offline'); }
    }, 2500);
    return () => clearTimeout(pushTimer.current);
  }, [history, settings, bodyweightLog, videos]);

  // --- push photos.json on change (debounced; large file, kept separate) ---
  useEffect(() => {
    if (!syncReadyRef.current || !sync.isConfigured()) return;
    clearTimeout(photoTimer.current);
    photoTimer.current = setTimeout(async () => {
      try { shaRef.current.photos = await sync.pushPhotos({ photos }, shaRef.current.photos); } catch (e) { /* retried on next change */ }
    }, 3000);
    return () => clearTimeout(photoTimer.current);
  }, [photos]);

  const pullMergePush = async () => {
    const m = await sync.pullMerge({ settings, history, photos, videos, bodyweightLog, updatedAt: localStorage.getItem('hyp_updatedAt') || '', hasLocal: Object.keys(history).length > 0 });
    shaRef.current = m.shas;
    setSettings(prev => ({ ...prev, ...m.settings })); setHistory(m.history); setBodyweightLog(m.bodyweightLog); setVideos(m.videos); setPhotos(m.photos);
    shaRef.current.history = await sync.pushHistory({ settings: m.settings, history: m.history, bodyweightLog: m.bodyweightLog, videos: m.videos }, m.shas.history);
    shaRef.current.photos = await sync.pushPhotos({ photos: m.photos }, m.shas.photos);
    localStorage.setItem('hyp_updatedAt', new Date().toISOString());
  };

  const saveSyncConfig = async (token, repo) => {
    sync.setConfig(token, repo);
    if (!token) { setSyncStatus('off'); return { ok: false, msg: 'Token cleared — sync off.' }; }
    setSyncStatus('syncing');
    const res = await sync.testConnection();
    if (!res.ok) { setSyncStatus('error'); return { ok: false, msg: res.status === 404 ? 'Repo not found — check the name and that the token can see it.' : `Failed (HTTP ${res.status}) — check the token.` }; }
    if (!res.canWrite) { setSyncStatus('error'); return { ok: false, msg: 'Token works but has no write access — set Contents: Read & Write.' }; }
    syncReadyRef.current = false;
    try { await pullMergePush(); setSyncStatus('synced'); setLastSyncAt(new Date().toISOString()); setTimeout(() => { syncReadyRef.current = true; }, 0); return { ok: true, msg: '✓ Connected and synced.' }; }
    catch (e) { setSyncStatus('error'); return { ok: false, msg: 'Connected but sync failed: ' + e.message }; }
  };

  const syncNow = async () => {
    if (!sync.isConfigured()) return;
    setSyncStatus('syncing');
    try { await pullMergePush(); setSyncStatus('synced'); setLastSyncAt(new Date().toISOString()); }
    catch (e) { setSyncStatus(navigator.onLine ? 'error' : 'offline'); }
  };

  useEffect(() => {
    const sessionActive = sessionStart && !sessionEnd;
    const restActive = restEndAt !== null;
    if (!sessionActive && !restActive) return;
    const id = setInterval(() => {
      setNowTick(t => t + 1);
      if (restEndAt && Date.now() >= restEndAt && !restNotifiedRef.current) {
        restNotifiedRef.current = true;
        if (navigator.vibrate) navigator.vibrate([200, 100, 200, 100, 200]);
      }
    }, 1000);
    return () => clearInterval(id);
  }, [sessionStart, sessionEnd, restEndAt]);

  useEffect(() => {
    const refresh = () => { if (!document.hidden) setNowTick(t => t + 1); };
    document.addEventListener('visibilitychange', refresh);
    window.addEventListener('focus', refresh);
    window.addEventListener('pageshow', refresh);
    return () => {
      document.removeEventListener('visibilitychange', refresh);
      window.removeEventListener('focus', refresh);
      window.removeEventListener('pageshow', refresh);
    };
  }, []);

  useEffect(() => {
    if (toast) {
      const t = setTimeout(() => setToast(null), 2500);
      return () => clearTimeout(t);
    }
  }, [toast]);

  const startRest = (seconds) => {
    setRestEndAt(Date.now() + seconds * 1000);
    setRestTotalSec(seconds);
    restNotifiedRef.current = false;
  };
  const stopRest = () => {
    setRestEndAt(null);
    setRestTotalSec(0);
    restNotifiedRef.current = true;
    window.storage.delete('hyp_rest').catch(() => {});
  };

  const startSession = () => {
    const k = sessionKey(todayKey(), selectedDay);
    const now = new Date().toISOString();
    setHistory(prev => {
      const rec = prev[k] || { date: todayKey(), day: selectedDay, sets: {} };
      return { ...prev, [k]: { ...rec, startTime: rec.startTime || now, finishedAt: null } };
    });
  };
  const endSession = () => {
    const k = sessionKey(todayKey(), selectedDay);
    setHistory(prev => {
      const rec = prev[k];
      if (!rec) return prev;
      return { ...prev, [k]: { ...rec, finishedAt: new Date().toISOString() } };
    });
    autoBackupToClipboard();
  };
  // Re-open a finished session (e.g. to add a forgotten set).
  const resetSession = () => {
    const k = sessionKey(todayKey(), selectedDay);
    setHistory(prev => {
      const rec = prev[k];
      if (!rec) return prev;
      return { ...prev, [k]: { ...rec, finishedAt: null } };
    });
  };

  const logSet = (exerciseId, exerciseName, totalSets, setData) => {
    const k = sessionKey(todayKey(), selectedDay);
    const now = new Date().toISOString();
    let newSetNumber = 1;
    setHistory(prev => {
      const rec = prev[k] || { date: todayKey(), day: selectedDay, sets: {}, startTime: now };
      const sets = rec.sets[exerciseId] || [];
      newSetNumber = sets.length + 1;
      return {
        ...prev,
        [k]: {
          ...rec,
          startTime: rec.startTime || now,
          finishedAt: null,
          sets: { ...rec.sets, [exerciseId]: [...sets, { ...setData, timestamp: now }] },
          endTime: now,
        },
      };
    });
    setToast({ exerciseName, setNumber: newSetNumber, totalSets, type: 'log' });
  };

  const undoLastSet = (exerciseId) => {
    const k = sessionKey(todayKey(), selectedDay);
    setHistory(prev => {
      const rec = prev[k];
      if (!rec || !rec.sets[exerciseId]) return prev;
      const sets = [...rec.sets[exerciseId]];
      sets.pop();
      return { ...prev, [k]: { ...rec, sets: { ...rec.sets, [exerciseId]: sets } } };
    });
  };

  const setPhoto = (exId, url) => setPhotos(p => ({ ...p, [exId]: url }));
  const setVideoUrl = (exId, url) => setVideos(v => ({ ...v, [exId]: url }));

  const buildBackupJSON = () => JSON.stringify({
    version: 5,
    exportedAt: new Date().toISOString(),
    settings, history, photos, videos, bodyweightLog,
  });

  const autoBackupToClipboard = async () => {
    try {
      const json = buildBackupJSON();
      await navigator.clipboard.writeText(json);
      const now = new Date().toISOString();
      setLastBackup(now);
      window.storage.set('hyp_last_backup', JSON.stringify(now)).catch(() => {});
      setToast({ type: 'backup', message: 'Backup copied to clipboard — paste it to Notes/email to save!' });
    } catch (e) {
      // Clipboard not available; fall back silently
    }
  };

  const manualBackupNow = async () => {
    try {
      const json = buildBackupJSON();
      await navigator.clipboard.writeText(json);
      const now = new Date().toISOString();
      setLastBackup(now);
      window.storage.set('hyp_last_backup', JSON.stringify(now)).catch(() => {});
      setToast({ type: 'backup', message: '✓ Backup copied. Paste somewhere safe.' });
    } catch (e) {
      // Fallback: trigger download
      exportJSON();
    }
  };

  const restoreFromClipboard = async () => {
    try {
      const text = await navigator.clipboard.readText();
      const data = JSON.parse(text);
      if (!data.version || !data.history) { alert('Invalid backup data in clipboard'); return; }
      if (!window.confirm(`Restore backup from ${new Date(data.exportedAt).toLocaleString()}? This will REPLACE current data.`)) return;
      if (data.settings) setSettings(s => ({ ...s, ...data.settings }));
      if (data.history) setHistory(migrateHistory(data.history));
      if (data.photos) setPhotos(data.photos);
      if (data.videos) setVideos(data.videos);
      if (data.bodyweightLog) setBodyweightLog(data.bodyweightLog);
      alert(`✓ Restored ${Object.keys(data.history).length} sessions from clipboard`);
    } catch (e) {
      alert('Could not restore from clipboard. Make sure you copied a valid backup JSON.');
    }
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    const programRows = [['Day', 'Exercise', 'Target', 'Sets', 'Rep Range', 'RIR', 'Tempo', 'Rest (s)', 'Cue']];
    DAYS_ORDER.forEach(d => {
      const day = PROGRAM[d];
      day.exercises.forEach(ex => programRows.push([day.label, ex.name, ex.target, ex.sets, `${ex.repMin}-${ex.repMax}`, `${ex.rirMin}-${ex.rirMax}`, ex.tempo, ex.rest, ex.cue]));
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(programRows), 'Program');
    const historyRows = [['Date', 'Day', 'Exercise', 'Set #', 'Weight', 'Unit', 'Per Side?', 'Reps', 'RIR', 'Time']];
    Object.entries(history).sort().forEach(([key, day]) => {
      const date = day.date || key.split('__')[0];
      Object.entries(day.sets || {}).forEach(([exId, sets]) => {
        const ex = Object.values(PROGRAM).flatMap(d => d.exercises || []).find(e => e.id === exId);
        sets.forEach((s, i) => historyRows.push([date, PROGRAM[day.day]?.label || day.day, ex?.name || exId, i + 1, s.weight, s.unit, s.weightMode === 'perSide' ? 'Yes' : 'No', s.reps, s.rir, s.timestamp]));
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(historyRows), 'History');
    const bwRows = [['Date', 'Bodyweight', 'Unit']];
    bodyweightLog.forEach(b => bwRows.push([b.date, b.weight, b.unit]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bwRows), 'Bodyweight');
    XLSX.writeFile(wb, `hypertrophy_log_${todayKey()}.xlsx`);
  };

  const exportJSON = () => {
    const json = buildBackupJSON();
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `hypertrophy_backup_${todayKey()}.json`; a.click();
    URL.revokeObjectURL(url);
    const now = new Date().toISOString();
    setLastBackup(now);
    window.storage.set('hyp_last_backup', JSON.stringify(now)).catch(() => {});
  };

  const importJSON = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = JSON.parse(e.target.result);
        if (data.settings) setSettings(s => ({ ...s, ...data.settings }));
        if (data.history) setHistory(migrateHistory(data.history));
        if (data.photos) setPhotos(data.photos);
        if (data.videos) setVideos(data.videos);
        if (data.bodyweightLog) setBodyweightLog(data.bodyweightLog);
        alert(`✓ Restored from ${data.exportedAt || 'backup'}`);
      } catch (err) { alert('Invalid backup file'); }
    };
    reader.readAsText(file);
  };

  const importExcel = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
        const allEx = Object.values(PROGRAM).flatMap(d => d.exercises || []);
        const nameToId = {};
        allEx.forEach(ex => { nameToId[ex.name.toLowerCase().trim()] = ex.id; });
        const labelToKey = {};
        Object.entries(PROGRAM).forEach(([key, day]) => { if (day.label) labelToKey[day.label.toLowerCase().trim()] = key; });
        const newHistory = {};
        let importedRows = 0, skippedRows = 0;
        const histSheet = wb.Sheets['History'];
        if (histSheet) {
          XLSX.utils.sheet_to_json(histSheet).forEach(row => {
            const date = String(row['Date'] || '').trim();
            const dayLabel = String(row['Day'] || '').toLowerCase().trim();
            const exName = String(row['Exercise'] || '').toLowerCase().trim();
            const weight = parseFloat(row['Weight']);
            const reps = parseInt(row['Reps']);
            if (!date || !dayLabel || !exName || isNaN(weight) || isNaN(reps)) { skippedRows++; return; }
            const dayKey = labelToKey[dayLabel];
            const exId = nameToId[exName];
            if (!dayKey || !exId) { skippedRows++; return; }
            const rir = parseInt(row['RIR']);
            const unit = String(row['Unit'] || 'kg').toLowerCase();
            const perSide = String(row['Per Side?'] || '').toLowerCase() === 'yes';
            let timestamp = row['Time'];
            if (!timestamp || isNaN(Date.parse(String(timestamp)))) timestamp = `${date}T12:00:00.000Z`;
            timestamp = String(timestamp);
            if (!newHistory[date]) newHistory[date] = { day: dayKey, sets: {}, startTime: timestamp };
            if (!newHistory[date].sets[exId]) newHistory[date].sets[exId] = [];
            newHistory[date].sets[exId].push({ weight, unit, reps, rir: isNaN(rir) ? 0 : rir, weightMode: perSide ? 'perSide' : 'total', timestamp });
            newHistory[date].endTime = timestamp;
            importedRows++;
          });
        }
        const newBw = [];
        const bwSheet = wb.Sheets['Bodyweight'];
        if (bwSheet) {
          XLSX.utils.sheet_to_json(bwSheet).forEach(row => {
            const date = String(row['Date'] || '').trim();
            const w = parseFloat(row['Bodyweight']);
            if (!date || isNaN(w)) return;
            newBw.push({ date, weight: w, unit: String(row['Unit'] || 'kg').toLowerCase() });
          });
        }
        const sessionCount = Object.keys(newHistory).length;
        const msg = `Import ${importedRows} sets across ${sessionCount} sessions + ${newBw.length} bodyweight entries?\n\nThis REPLACES current data.\n${skippedRows > 0 ? `\n${skippedRows} rows skipped (unrecognized names).` : ''}\n\nContinue?`;
        if (!window.confirm(msg)) return;
        setHistory(migrateHistory(newHistory));
        setBodyweightLog(newBw);
        alert(`✓ Imported ${importedRows} sets`);
      } catch (err) { alert('Could not read Excel file.'); }
    };
    reader.readAsArrayBuffer(file);
  };

  const sessionElapsed = sessionStart ? Math.floor(((sessionEnd ? new Date(sessionEnd) : new Date()) - new Date(sessionStart)) / 1000) : 0;
  const currentDayData = PROGRAM[selectedDay];
  const todayHistory = history[currentKey] || { sets: {} };

  // Data statistics for visible status banner
  const totalSessions = Object.keys(history).length;
  const totalSets = Object.values(history).reduce((sum, day) => sum + Object.values(day.sets || {}).reduce((s, sets) => s + sets.length, 0), 0);
  const totalPhotos = Object.keys(photos).length;
  const daysSinceBackup = lastBackup ? Math.floor((Date.now() - new Date(lastBackup)) / 86400000) : null;

  return (
    <div style={{ fontFamily: "'DM Sans', system-ui, sans-serif", minHeight: '100vh', background: '#0a0a0a', color: '#f5f5f5', paddingBottom: '90px' }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;700&display=swap');
        * { box-sizing: border-box; -webkit-tap-highlight-color: transparent; }
        body { margin: 0; }
        .display { font-family: 'Bebas Neue', sans-serif; letter-spacing: 0.02em; }
        .mono { font-family: 'JetBrains Mono', monospace; }
        button { font-family: inherit; cursor: pointer; border: none; }
        input { font-family: inherit; }
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-thumb { background: #333; border-radius: 3px; }
        @keyframes slideIn { from { transform: translateY(100%); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        .toast-anim { animation: slideIn 0.3s ease-out; }
      `}</style>

      {/* HEADER */}
      <div style={{ position: 'sticky', top: 0, background: '#0a0a0a', borderBottom: '1px solid #1f1f1f', zIndex: 50, padding: '10px 14px 8px' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <img src={`${import.meta.env.BASE_URL}assets/logo.png`} alt="" style={{ width: 36, height: 36 }} onError={(e) => { e.target.style.display = 'none'; }} />
            <div className="display" style={{ fontSize: 22, lineHeight: 1, color: '#facc15' }}>STRONGER EVERY DAY</div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {sessionStart && !sessionEnd && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 5, background: '#22c55e', color: '#0a0a0a', padding: '4px 8px', borderRadius: 6, fontWeight: 700, lineHeight: 1.1 }}>
                <Clock size={11} />
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                  <span className="mono" style={{ fontSize: 12 }}>{formatTime(sessionElapsed)}</span>
                  <span style={{ fontSize: 8, opacity: 0.7 }}>since {formatClockTime(sessionStart)}</span>
                </div>
              </div>
            )}
            {syncStatus !== 'off' && (
              <button onClick={syncNow} title="Sync now" style={{ display: 'flex', alignItems: 'center', gap: 4, background: 'none', padding: '2px 4px' }}>
                <span style={{ width: 8, height: 8, borderRadius: 4, background: syncStatus === 'synced' ? '#22c55e' : syncStatus === 'syncing' ? '#facc15' : syncStatus === 'offline' ? '#888' : '#ef4444' }} />
                <span style={{ fontSize: 8, color: '#666', fontWeight: 700 }}>{syncStatus === 'syncing' ? 'SYNC…' : syncStatus === 'synced' ? 'SYNCED' : syncStatus === 'offline' ? 'OFFLINE' : syncStatus === 'error' ? 'SYNC ⚠' : 'SYNC'}</span>
              </button>
            )}
            <div className="mono" style={{ fontSize: 10, color: '#666' }}>WK {settings.mesoWeek}/6</div>
          </div>
        </div>
        {/* DATA STATUS BANNER */}
        <DataStatusBanner totalSessions={totalSessions} totalSets={totalSets} totalPhotos={totalPhotos} daysSinceBackup={daysSinceBackup} onBackupNow={manualBackupNow} onRestore={restoreFromClipboard} syncOn={syncStatus !== 'off'} />
        {view === 'today' && <DaySelector selected={selectedDay} onSelect={setSelectedDay} />}
      </div>

      {/* REST TIMER OVERLAY */}
      {restEndAt && (() => {
        const remaining = Math.max(0, Math.ceil((restEndAt - Date.now()) / 1000));
        const done = remaining === 0;
        return (
          <div style={{ position: 'fixed', top: 130, right: 14, zIndex: 100, background: done ? '#22c55e' : '#facc15', color: '#0a0a0a', padding: '10px 14px', borderRadius: 12, boxShadow: '0 8px 24px rgba(250,204,21,0.4)', display: 'flex', alignItems: 'center', gap: 8 }}>
            <Clock size={16} />
            <div>
              <div className="mono" style={{ fontSize: 20, fontWeight: 700, lineHeight: 1 }}>{done ? 'GO' : formatTime(remaining)}</div>
              <div style={{ fontSize: 8, opacity: 0.7, marginTop: 2 }}>{done ? 'REST DONE' : 'REST'}</div>
            </div>
            <button onClick={stopRest} style={{ background: 'rgba(0,0,0,0.15)', color: '#0a0a0a', padding: 5, borderRadius: 6 }}><X size={12} /></button>
          </div>
        );
      })()}

      {/* TOAST */}
      {toast && (
        <div className="toast-anim" style={{ position: 'fixed', bottom: 100, left: '50%', transform: 'translateX(-50%)', zIndex: 200, background: toast.type === 'backup' ? '#3b82f6' : '#22c55e', color: '#fff', padding: '10px 16px', borderRadius: 10, fontWeight: 700, fontSize: 12, display: 'flex', alignItems: 'center', gap: 8, boxShadow: '0 8px 24px rgba(0,0,0,0.4)', maxWidth: '90%' }}>
          <Check size={16} strokeWidth={3} />
          {toast.type === 'backup' ? <span>{toast.message}</span> : <span>SET {toast.setNumber} OF {toast.totalSets} LOGGED</span>}
        </div>
      )}

      {/* PHOTO ZOOM LIGHTBOX */}
      {zoomPhoto && (
        <div onClick={() => setZoomPhoto(null)} style={{ position: 'fixed', inset: 0, zIndex: 300, background: 'rgba(0,0,0,0.92)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16 }}>
          <img src={zoomPhoto} alt="" style={{ maxWidth: '100%', maxHeight: '100%', borderRadius: 10, objectFit: 'contain' }} />
          <button onClick={() => setZoomPhoto(null)} style={{ position: 'fixed', top: 16, right: 16, background: 'rgba(255,255,255,0.12)', color: '#fff', width: 38, height: 38, borderRadius: 19, display: 'flex', alignItems: 'center', justifyContent: 'center' }}><X size={20} /></button>
          <div style={{ position: 'fixed', bottom: 20, left: 0, right: 0, textAlign: 'center', color: '#888', fontSize: 11 }}>Tap anywhere to close</div>
        </div>
      )}

      {/* CONTENT */}
      <div style={{ padding: '12px 14px' }}>
        {view === 'today' && (
          <TodayView day={currentDayData} dayKey={selectedDay} settings={settings} history={history}
            todayHistory={todayHistory} photos={photos} videos={videos}
            onLogSet={logSet} onUndo={undoLastSet} onPhoto={setPhoto} onVideo={setVideoUrl}
            onStartRest={startRest} setSettings={setSettings} onZoom={setZoomPhoto}
            sessionStart={sessionStart} sessionEnd={sessionEnd}
            onStartSession={startSession} onEndSession={endSession} onResetSession={resetSession} />
        )}
        {view === 'history' && <HistoryView history={history} photos={photos} onZoom={setZoomPhoto} />}
        {view === 'progress' && <ProgressView history={history} bodyweightLog={bodyweightLog} setBodyweightLog={setBodyweightLog} settings={settings} />}
        {view === 'settings' && <SettingsView settings={settings} setSettings={setSettings} onExportExcel={exportToExcel} onExportJSON={exportJSON} onImportJSON={importJSON} onImportExcel={importExcel} onBackupNow={manualBackupNow} onRestore={restoreFromClipboard} lastBackup={lastBackup} syncStatus={syncStatus} lastSyncAt={lastSyncAt} onSaveSync={saveSyncConfig} onSyncNow={syncNow} />}
      </div>

      {/* BOTTOM NAV */}
      <div style={{ position: 'fixed', bottom: 0, left: 0, right: 0, background: '#0a0a0a', borderTop: '1px solid #1f1f1f', display: 'flex', justifyContent: 'space-around', padding: '10px 0 14px', zIndex: 50 }}>
        {[
          { id: 'today', icon: Dumbbell, label: 'Today' },
          { id: 'history', icon: Calendar, label: 'History' },
          { id: 'progress', icon: TrendingUp, label: 'Progress' },
          { id: 'settings', icon: Settings, label: 'Settings' },
        ].map(tab => (
          <button key={tab.id} onClick={() => setView(tab.id)} style={{ background: 'none', color: view === tab.id ? '#facc15' : '#666', display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4, padding: '4px 12px' }}>
            <tab.icon size={20} />
            <span style={{ fontSize: 10, fontWeight: 500 }}>{tab.label}</span>
          </button>
        ))}
      </div>
    </div>
  );
}

// ============================================================
// DATA STATUS BANNER — persistent at top of every screen
// ============================================================

function DataStatusBanner({ totalSessions, totalSets, totalPhotos, daysSinceBackup, onBackupNow, onRestore, syncOn }) {
  const isStale = daysSinceBackup === null || daysSinceBackup > 7;
  // When GitHub sync is on, backup is automatic — show a clean synced strip, no manual backup buttons.
  if (syncOn) {
    return (
      <div style={{ background: '#0f1a14', border: '1px solid #1f3a2c', borderRadius: 8, padding: '6px 10px', marginBottom: 8, display: 'flex', alignItems: 'center', gap: 8 }}>
        <Database size={12} color="#22c55e" />
        <span className="mono" style={{ fontSize: 10, color: '#ddd', fontWeight: 700 }}>{totalSessions} sessions · {totalSets} sets · {totalPhotos} photos</span>
        <span style={{ marginLeft: 'auto', fontSize: 9, color: '#6ee7b7', fontWeight: 700 }}>☁ Auto-synced to GitHub</span>
      </div>
    );
  }
  return (
    <div style={{ background: isStale ? '#1a0f00' : '#0f1a0f', border: `1px solid ${isStale ? '#facc15' : '#22c55e'}`, borderRadius: 8, padding: '6px 10px', marginBottom: 8, display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flex: 1, minWidth: 0 }}>
        <Database size={12} color={isStale ? '#facc15' : '#22c55e'} />
        <div style={{ display: 'flex', flexDirection: 'column', gap: 1 }}>
          <span className="mono" style={{ fontSize: 10, color: '#ddd', fontWeight: 700 }}>{totalSessions} sessions · {totalSets} sets · {totalPhotos} photos</span>
          <span style={{ fontSize: 9, color: isStale ? '#facc15' : '#22c55e' }}>
            {daysSinceBackup === null ? 'NO BACKUP YET — backup now!' : daysSinceBackup === 0 ? '✓ Backed up today' : `Last backup: ${daysSinceBackup}d ago${daysSinceBackup > 7 ? ' — STALE' : ''}`}
          </span>
        </div>
      </div>
      <div style={{ display: 'flex', gap: 4 }}>
        <button onClick={onBackupNow} style={{ background: '#facc15', color: '#0a0a0a', padding: '4px 8px', borderRadius: 5, fontSize: 9, fontWeight: 700, display: 'flex', alignItems: 'center', gap: 3 }}><Copy size={9} /> BACKUP</button>
        <button onClick={onRestore} style={{ background: '#1a1a1a', color: '#facc15', padding: '4px 8px', borderRadius: 5, fontSize: 9, fontWeight: 700, border: '1px solid #2a2a2a', display: 'flex', alignItems: 'center', gap: 3 }}><ClipboardPaste size={9} /> RESTORE</button>
      </div>
    </div>
  );
}

// ============================================================
// DAY SELECTOR
// ============================================================

function DaySelector({ selected, onSelect }) {
  return (
    <div style={{ display: 'flex', gap: 5, overflowX: 'auto' }}>
      {DAYS_ORDER.map(d => {
        const dayData = PROGRAM[d];
        const isSelected = d === selected;
        return (
          <button key={d} onClick={() => onSelect(d)} style={{
            background: isSelected ? '#facc15' : '#1a1a1a',
            color: isSelected ? '#0a0a0a' : '#ccc',
            border: isSelected ? 'none' : '1px solid #2a2a2a',
            padding: '7px 11px', borderRadius: 8, fontWeight: 700, fontSize: 11, minWidth: 56, whiteSpace: 'nowrap',
          }}>{dayData.label}</button>
        );
      })}
    </div>
  );
}

// ============================================================
// TODAY VIEW
// ============================================================

function TodayView({ day, dayKey, settings, history, todayHistory, photos, videos, onLogSet, onUndo, onPhoto, onVideo, onStartRest, setSettings, onZoom, sessionStart, sessionEnd, onStartSession, onEndSession, onResetSession }) {
  const [warmupOpen, setWarmupOpen] = useState(false);
  const [rirInfoOpen, setRirInfoOpen] = useState(false);

  return (
    <div>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 12 }}>
        <div>
          <div className="display" style={{ fontSize: 30, lineHeight: 1, color: '#facc15' }}>{day.label}: {day.name}</div>
          <div style={{ color: '#888', fontSize: 12, marginTop: 4 }}>{day.focus}</div>
        </div>
        <SessionControl sessionStart={sessionStart} sessionEnd={sessionEnd} onStart={onStartSession} onEnd={onEndSession} onReset={onResetSession} />
      </div>

      <button onClick={() => setRirInfoOpen(!rirInfoOpen)} style={{ width: '100%', background: '#161616', border: '1px solid #2a2a2a', padding: '8px 10px', borderRadius: 8, color: '#facc15', fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', gap: 6, marginBottom: 8 }}>
        <Info size={12} />
        <span>WHAT DO THE 0,1,2,3 BUTTONS MEAN?</span>
        {rirInfoOpen ? <ChevronDown size={12} style={{ marginLeft: 'auto' }} /> : <ChevronRight size={12} style={{ marginLeft: 'auto' }} />}
      </button>
      {rirInfoOpen && (
        <div style={{ background: '#161616', border: '1px solid #2a2a2a', padding: 10, borderRadius: 8, marginBottom: 8, fontSize: 11, color: '#ccc', lineHeight: 1.5 }}>
          <div style={{ marginBottom: 6 }}>The <strong style={{ color: '#facc15' }}>0/1/2/3</strong> buttons are <strong>RIR (Reps in Reserve)</strong>. After each set, tap how many MORE reps you could have done.</div>
          <div style={{ display: 'grid', gridTemplateColumns: '28px 1fr', gap: 5, fontSize: 10.5 }}>
            <span className="mono" style={{ background: '#facc15', color: '#0a0a0a', textAlign: 'center', borderRadius: 4, fontWeight: 700 }}>0</span><span>Failure</span>
            <span className="mono" style={{ background: '#facc15', color: '#0a0a0a', textAlign: 'center', borderRadius: 4, fontWeight: 700 }}>1</span><span>1 more rep possible</span>
            <span className="mono" style={{ background: '#facc15', color: '#0a0a0a', textAlign: 'center', borderRadius: 4, fontWeight: 700 }}>2</span><span>2 more reps possible</span>
            <span className="mono" style={{ background: '#facc15', color: '#0a0a0a', textAlign: 'center', borderRadius: 4, fontWeight: 700 }}>3</span><span>3 more — too easy</span>
          </div>
        </div>
      )}

      <CollapsibleSection open={warmupOpen} setOpen={setWarmupOpen} icon={<Flame size={14} color="#facc15" />} title="WARMUP (8–10 MIN)">
        <WarmupContent mobility={day.mobility} />
      </CollapsibleSection>

      <div style={{ marginTop: 12 }}>
        {day.exercises.map((ex, idx) => (
          <ExerciseCard key={ex.id + idx} index={idx + 1} exercise={ex} settings={settings} history={history}
            todaysets={todayHistory.sets[ex.id] || []} photo={photos[ex.id]} customVideo={videos[ex.id]}
            onLogSet={(setData) => onLogSet(ex.id, ex.name, ex.sets, setData)} onUndo={() => onUndo(ex.id)}
            onPhoto={(url) => onPhoto(ex.id, url)} onVideoUpdate={(url) => onVideo(ex.id, url)} onZoom={onZoom}
            onStartRest={() => onStartRest(ex.rest)} />
        ))}
      </div>
    </div>
  );
}

function SessionControl({ sessionStart, sessionEnd, onStart, onEnd, onReset }) {
  if (sessionStart && !sessionEnd) return <button onClick={onEnd} style={{ background: '#ef4444', color: '#fff', padding: '7px 11px', borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', gap: 5 }}><Square size={12} /> END</button>;
  if (sessionEnd) return <button onClick={onReset} style={{ background: '#1a1a1a', color: '#888', padding: '7px 11px', borderRadius: 7, fontWeight: 700, fontSize: 11, border: '1px solid #2a2a2a' }}>DONE ✓</button>;
  return <button onClick={onStart} style={{ background: '#22c55e', color: '#0a0a0a', padding: '7px 11px', borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', gap: 5 }}><Play size={12} /> START</button>;
}

function CollapsibleSection({ open, setOpen, icon, title, children }) {
  return (
    <div style={{ marginBottom: 8, background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, overflow: 'hidden' }}>
      <button onClick={() => setOpen(!open)} style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '10px 12px', background: 'none', color: '#f5f5f5' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>{icon}<span style={{ fontSize: 11, fontWeight: 700, letterSpacing: '0.05em' }}>{title}</span></div>
        {open ? <ChevronDown size={14} color="#666" /> : <ChevronRight size={14} color="#666" />}
      </button>
      {open && <div style={{ padding: '0 12px 12px' }}>{children}</div>}
    </div>
  );
}

function WarmupContent({ mobility }) {
  const [checked, setChecked] = useState({});
  const items = [
    { label: '5 min light cardio — bike/rower at Zone 2', key: 'cardio' },
    ...mobility.map((m, i) => ({ label: m, key: `mob_${i}` })),
    { label: 'Specific warmup: 50%×8, 70%×5, 85%×3 (no failure)', key: 'specific' },
  ];
  return (
    <div style={{ fontSize: 12 }}>
      {items.map(item => (
        <button key={item.key} onClick={() => setChecked(c => ({ ...c, [item.key]: !c[item.key] }))} style={{ width: '100%', display: 'flex', alignItems: 'center', gap: 10, padding: '7px 0', background: 'none', color: checked[item.key] ? '#666' : '#ddd', textAlign: 'left', textDecoration: checked[item.key] ? 'line-through' : 'none', fontSize: 12 }}>
          <div style={{ width: 16, height: 16, borderRadius: 4, border: '2px solid', borderColor: checked[item.key] ? '#facc15' : '#444', background: checked[item.key] ? '#facc15' : 'transparent', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
            {checked[item.key] && <Check size={10} color="#0a0a0a" strokeWidth={3} />}
          </div>
          <span>{item.label}</span>
        </button>
      ))}
    </div>
  );
}

// ============================================================
// EXERCISE CARD
// ============================================================

function ExerciseCard({ index, exercise, settings, history, todaysets, photo, customVideo, onLogSet, onUndo, onPhoto, onVideoUpdate, onStartRest, onZoom }) {
  const [expanded, setExpanded] = useState(false);
  const [editingVideo, setEditingVideo] = useState(false);
  const [videoInput, setVideoInput] = useState('');
  const [weight, setWeight] = useState('');
  const [reps, setReps] = useState('');
  const [rir, setRir] = useState(1);
  const fileRef = useRef(null);

  const lastSession = (() => {
    const keys = Object.keys(history).sort().reverse();
    for (const k of keys) {
      if (history[k]?.date === todayKey()) continue; // skip today's sessions
      if (history[k].sets?.[exercise.id]?.length) return { date: history[k].date || k.split('__')[0], sets: history[k].sets[exercise.id] };
    }
    return null;
  })();

  const suggestion = suggestWeight(lastSession, exercise.repMin, exercise.repMax, exercise.rirMin);

  // Pre-fill the weight: from your last logged session if available, otherwise
  // from the program's seeded starting weight (your real numbers).
  useEffect(() => {
    if (weight !== '' || reps !== '') return;
    if (suggestion) {
      const w = suggestion.unit === settings.unit ? suggestion.weight : (suggestion.unit === 'kg' ? kgToLb(suggestion.weight) : lbToKg(suggestion.weight));
      setWeight(w); setReps(exercise.repMin); setRir(exercise.rirMin);
    } else if (exercise.startWeight != null) {
      setWeight(exercise.startWeight || ''); setReps(exercise.repMin); setRir(exercise.rirMin);
    }
  }, [lastSession?.date, exercise.id]);

  const completedSets = todaysets.length;
  const currentSetNumber = Math.min(completedSets + 1, exercise.sets);
  const isComplete = completedSets >= exercise.sets;
  const details = EXERCISE_DETAILS[exercise.id];
  const videoUrl = customVideo || `https://www.youtube.com/results?search_query=${encodeURIComponent(exercise.name + ' jeff nippard form')}`;

  const handlePhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    onPhoto(await compressImage(file));
  };

  const handleLog = () => {
    if (!weight || !reps) return;
    onLogSet({ weight: parseFloat(weight), unit: settings.unit, reps: parseInt(reps), rir: parseInt(rir), weightMode: exercise.weightMode });
    onStartRest();
  };

  const saveVideo = () => {
    onVideoUpdate(videoInput.trim() || '');
    setEditingVideo(false);
    setVideoInput('');
  };

  const conversionDisplay = weight ? (settings.unit === 'kg' ? kgToLb(parseFloat(weight)) : lbToKg(parseFloat(weight))) : null;

  return (
    <div style={{ background: '#111', border: `1px solid ${isComplete ? '#22c55e' : '#1f1f1f'}`, borderRadius: 12, marginBottom: 10, overflow: 'hidden' }}>
      <div style={{ display: 'flex', alignItems: 'flex-start', padding: 11, gap: 10 }}>
        <div style={{ position: 'relative', width: 54, height: 54, flexShrink: 0 }}>
          <button onClick={() => (photo ? onZoom?.(photo) : fileRef.current?.click())} title={photo ? 'Tap to zoom' : 'Add photo'} style={{ width: 54, height: 54, borderRadius: 9, background: photo ? `url(${photo}) center/cover` : '#1a1a1a', border: '1px solid #2a2a2a', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 0 }}>
            {!photo && <Camera size={16} color="#555" />}
            {photo && <span style={{ position: 'absolute', bottom: 2, left: 2, background: 'rgba(0,0,0,0.55)', borderRadius: 4, padding: '1px 3px', fontSize: 8, color: '#fff', fontWeight: 700, letterSpacing: '0.05em' }}>🔍</span>}
          </button>
          {photo && <button onClick={() => fileRef.current?.click()} title="Change photo" style={{ position: 'absolute', bottom: -5, right: -5, width: 20, height: 20, borderRadius: 10, background: '#facc15', border: '2px solid #111', display: 'flex', alignItems: 'center', justifyContent: 'center' }}><Camera size={9} color="#0a0a0a" /></button>}
          <input ref={fileRef} type="file" accept="image/*" onChange={handlePhoto} style={{ display: 'none' }} />
        </div>
        <div style={{ flex: 1, minWidth: 0 }}>
          <div style={{ display: 'flex', alignItems: 'baseline', gap: 5, marginBottom: 2 }}>
            <span className="mono" style={{ fontSize: 9, color: '#666', fontWeight: 700 }}>#{index}</span>
            <div style={{ fontSize: 13, fontWeight: 700, color: '#f5f5f5', lineHeight: 1.2 }}>{exercise.name}</div>
          </div>
          <div style={{ fontSize: 11, color: '#888', display: 'flex', alignItems: 'center', gap: 4, marginBottom: 4 }}>
            <Target size={9} /><span>{exercise.target}</span>
          </div>
          {exercise.garmin && <div className="mono" style={{ fontSize: 9, color: '#3b82f6', marginBottom: 6 }}>Garmin: {exercise.garmin}</div>}
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexWrap: 'wrap' }}>
            <div style={{ background: isComplete ? '#22c55e' : '#facc15', color: '#0a0a0a', padding: '3px 8px', borderRadius: 5, fontSize: 10, fontWeight: 700 }}>
              {isComplete ? `✓ ${exercise.sets}/${exercise.sets} DONE` : `SET ${currentSetNumber}/${exercise.sets}`}
            </div>
            <Tag>{exercise.repMin}-{exercise.repMax} reps</Tag>
            <Tag>⏱{exercise.rest}s</Tag>
          </div>
        </div>
      </div>

      {lastSession && (
        <div style={{ padding: '0 11px 8px' }}>
          <div style={{ background: '#161616', borderRadius: 6, padding: '6px 9px', fontSize: 10, lineHeight: 1.5 }}>
            <div style={{ color: '#888' }}>
              <span style={{ color: '#666' }}>LAST ({lastSession.date}):</span>{' '}
              <span className="mono" style={{ color: '#ccc' }}>{lastSession.sets.map(s => `${s.weight}${s.weightMode === 'perSide' ? '×2' : ''}${s.unit}×${s.reps}`).join(' · ')}</span>
            </div>
            {suggestion && <div style={{ color: '#facc15', fontWeight: 700, marginTop: 2 }}>→ Try: {suggestion.weight} {suggestion.unit} ({suggestion.reason})</div>}
          </div>
        </div>
      )}

      <div style={{ height: 3, background: '#1a1a1a' }}>
        <div style={{ height: '100%', width: `${(completedSets / exercise.sets) * 100}%`, background: isComplete ? '#22c55e' : '#facc15', transition: 'width 0.3s' }} />
      </div>

      {!isComplete && (
        <>
          <div style={{ padding: '9px 10px 0' }}>
            <span style={{ display: 'inline-block', background: exercise.weightMode === 'perSide' ? '#3a1d0a' : '#0f1a14', color: exercise.weightMode === 'perSide' ? '#fdba74' : '#6ee7b7', border: `1px solid ${exercise.weightMode === 'perSide' ? '#7c3a12' : '#1f3a2c'}`, padding: '2px 8px', borderRadius: 5, fontSize: 9, fontWeight: 700, letterSpacing: '0.03em' }}>
              {exercise.weightMode === 'perSide' ? '⚖ LOG PER SIDE — one arm / one cable' : '⚖ LOG TOTAL — the machine/stack number'}
            </span>
          </div>
          <div style={{ padding: 10, display: 'flex', alignItems: 'center', gap: 5 }}>
            <NumInput value={weight} onChange={setWeight} placeholder={settings.unit} step={2.5} hint={exercise.weightMode === 'perSide' ? '/side' : ''} />
            <span style={{ color: '#444', fontSize: 13 }}>×</span>
            <NumInput value={reps} onChange={setReps} placeholder="reps" step={1} small />
            <button onClick={handleLog} disabled={!weight || !reps} style={{ background: weight && reps ? '#facc15' : '#1a1a1a', color: weight && reps ? '#0a0a0a' : '#444', padding: '9px 11px', borderRadius: 7, fontWeight: 700, fontSize: 11 }}>LOG</button>
          </div>
          <div style={{ padding: '0 10px 7px', display: 'flex', alignItems: 'center', gap: 7 }}>
            <span style={{ fontSize: 9, color: '#888', fontWeight: 700, letterSpacing: '0.05em' }}>RIR:</span>
            <RIRSelect value={rir} onChange={setRir} />
            <span style={{ fontSize: 9, color: '#555' }}>(reps left)</span>
          </div>
          {weight && (
            <div style={{ padding: '0 10px 7px', fontSize: 10, color: '#666', display: 'flex', justifyContent: 'space-between' }}>
              <span className="mono">= {conversionDisplay} {settings.unit === 'kg' ? 'lb' : 'kg'}</span>
              {exercise.weightMode === 'perSide' && <span>Total: {parseFloat(weight) * 2} {settings.unit}</span>}
            </div>
          )}
        </>
      )}

      {todaysets.length > 0 && (
        <div style={{ padding: '0 11px 9px' }}>
          {todaysets.map((s, i) => (
            <div key={i} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '5px 9px', background: '#0a0a0a', borderRadius: 5, marginBottom: 3 }}>
              <div className="mono" style={{ fontSize: 10, color: '#ddd' }}>
                <span style={{ color: '#22c55e' }}>SET {i + 1}/{exercise.sets} ✓</span> · {s.weight}{s.weightMode === 'perSide' ? '×2' : ''} {s.unit} · {s.reps}r · RIR {s.rir}
              </div>
              {i === todaysets.length - 1 && <button onClick={onUndo} style={{ background: 'none', color: '#666', padding: 3 }}><RotateCcw size={10} /></button>}
            </div>
          ))}
        </div>
      )}

      <button onClick={() => setExpanded(!expanded)} style={{ width: '100%', padding: 9, background: '#0d0d0d', color: '#facc15', fontSize: 10, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4, borderTop: '1px solid #1a1a1a' }}>
        {expanded ? 'COLLAPSE' : '📖 FULL EXERCISE GUIDE'} {expanded ? <ChevronDown size={11} /> : <ChevronRight size={11} />}
      </button>

      {expanded && details && <ExpandedDetails exercise={exercise} details={details} customVideo={customVideo} videoUrl={videoUrl} editingVideo={editingVideo} videoInput={videoInput} setEditingVideo={setEditingVideo} setVideoInput={setVideoInput} saveVideo={saveVideo} />}
    </div>
  );
}

function ExpandedDetails({ exercise, details, customVideo, videoUrl, editingVideo, videoInput, setEditingVideo, setVideoInput, saveVideo }) {
  return (
    <div style={{ padding: 12, borderTop: '1px solid #1a1a1a', background: '#0d0d0d' }}>
      <DetailSection icon="⚡" color="#facc15" title="QUICK CUE">
        <div style={{ fontSize: 12, color: '#ddd', lineHeight: 1.5 }}>{exercise.cue}</div>
      </DetailSection>
      <DetailSection icon="⏱️" color="#3b82f6" title="TEMPO">
        <div style={{ fontSize: 12, color: '#ddd' }}>
          <span className="mono" style={{ color: '#3b82f6', fontWeight: 700, fontSize: 13 }}>{exercise.tempo}</span>
          <div style={{ fontSize: 10.5, color: '#888', marginTop: 3 }}>{parseTempo(exercise.tempo)}</div>
        </div>
      </DetailSection>
      <DetailSection icon="🎯" color="#22c55e" title="SETUP">
        {details.setup.map((item, i) => (
          <div key={i} style={{ fontSize: 11.5, color: '#ddd', lineHeight: 1.5, marginBottom: 5, paddingLeft: 14, position: 'relative' }}>
            <span style={{ position: 'absolute', left: 0, color: '#22c55e' }}>▸</span>{item}
          </div>
        ))}
      </DetailSection>
      <DetailSection icon="▶️" color="#facc15" title="EXECUTION">
        {details.execution.map((item, i) => (
          <div key={i} style={{ fontSize: 11.5, color: '#ddd', lineHeight: 1.5, marginBottom: 7, paddingLeft: 22, position: 'relative' }}>
            <span style={{ position: 'absolute', left: 0, top: 0, background: '#facc15', color: '#0a0a0a', width: 16, height: 16, borderRadius: 3, fontSize: 9, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>{i + 1}</span>
            {item}
          </div>
        ))}
      </DetailSection>
      <DetailSection icon="💪" color="#a855f7" title="WHAT YOU SHOULD FEEL">
        {details.feel.map((item, i) => (
          <div key={i} style={{ fontSize: 11.5, color: '#ddd', lineHeight: 1.5, marginBottom: 5, paddingLeft: 14, position: 'relative' }}>
            <span style={{ position: 'absolute', left: 0, color: '#a855f7' }}>✓</span>{item}
          </div>
        ))}
      </DetailSection>
      <DetailSection icon="⚠️" color="#ef4444" title="COMMON MISTAKES">
        {details.mistakes.map((m, i) => (
          <div key={i} style={{ marginBottom: 6, padding: '7px 9px', background: '#1a0d0d', border: '1px solid #2a1a1a', borderRadius: 5 }}>
            <div style={{ fontSize: 11, color: '#fca5a5', marginBottom: 2, fontWeight: 700 }}>✗ {m.mistake}</div>
            <div style={{ fontSize: 10.5, color: '#bbb' }}><span style={{ color: '#22c55e', fontWeight: 700 }}>→</span> {m.fix}</div>
          </div>
        ))}
      </DetailSection>
      <DetailSection icon="📖" color="#888" title="FULL EXPLANATION">
        <div style={{ fontSize: 11.5, color: '#ccc', lineHeight: 1.6, fontStyle: 'italic' }}>{details.paragraph}</div>
      </DetailSection>
      <div style={{ marginTop: 12, paddingTop: 12, borderTop: '1px solid #1a1a1a' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 5 }}>
          <div style={{ fontSize: 9, fontWeight: 700, color: '#facc15', letterSpacing: '0.08em' }}>FORM VIDEO</div>
          <button onClick={() => { setEditingVideo(!editingVideo); setVideoInput(videoUrl); }} style={{ background: 'none', color: '#666', padding: 3, display: 'flex', alignItems: 'center', gap: 3, fontSize: 9 }}>
            <Edit3 size={9} /> {customVideo ? 'EDIT' : 'SET YOURS'}
          </button>
        </div>
        {editingVideo ? (
          <div style={{ display: 'flex', gap: 5 }}>
            <input type="url" value={videoInput} onChange={e => setVideoInput(e.target.value)} placeholder="Paste YouTube URL" style={{ flex: 1, background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 7, borderRadius: 5, fontSize: 10 }} />
            <button onClick={saveVideo} style={{ background: '#facc15', color: '#0a0a0a', padding: '0 10px', borderRadius: 5, fontWeight: 700, fontSize: 9 }}>SAVE</button>
          </div>
        ) : (
          <a href={videoUrl} target="_blank" rel="noopener noreferrer" style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, background: '#1a1a1a', color: '#facc15', padding: 9, borderRadius: 7, fontSize: 10, fontWeight: 700, border: '1px solid #2a2a2a', textDecoration: 'none' }}>
            <Youtube size={11} /> {customVideo ? 'WATCH YOUR VIDEO' : 'SEARCH FORM VIDEO'}
          </a>
        )}
      </div>
    </div>
  );
}

function DetailSection({ icon, color, title, children }) {
  return (
    <div style={{ marginBottom: 12 }}>
      <div style={{ fontSize: 9, fontWeight: 700, color, letterSpacing: '0.08em', marginBottom: 6, display: 'flex', alignItems: 'center', gap: 5 }}>
        <span>{icon}</span><span>{title}</span>
      </div>
      {children}
    </div>
  );
}

function parseTempo(t) {
  const [e, s, c, p] = t.split('-').map(Number);
  return `${e}s down · ${s}s stretch · ${c}s up · ${p}s squeeze`;
}

function Tag({ children }) {
  return <span style={{ fontSize: 9.5, color: '#aaa', background: '#1a1a1a', padding: '3px 6px', borderRadius: 4, fontWeight: 700 }}>{children}</span>;
}

function NumInput({ value, onChange, placeholder, step = 1, hint, small }) {
  return (
    <div style={{ position: 'relative', flex: small ? 0.6 : 1 }}>
      <input type="number" inputMode="decimal" value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder} step={step} style={{ width: '100%', background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: '9px 7px', borderRadius: 7, fontSize: 13, fontWeight: 700, fontFamily: 'JetBrains Mono, monospace', textAlign: 'center' }} />
      {hint && <span style={{ position: 'absolute', right: 5, top: '50%', transform: 'translateY(-50%)', fontSize: 8, color: '#555', pointerEvents: 'none' }}>{hint}</span>}
    </div>
  );
}

function RIRSelect({ value, onChange }) {
  return (
    <div style={{ display: 'flex', background: '#0a0a0a', borderRadius: 7, border: '1px solid #2a2a2a', padding: 2 }}>
      {[0, 1, 2, 3].map(n => (
        <button key={n} onClick={() => onChange(n)} style={{ background: value === n ? '#facc15' : 'transparent', color: value === n ? '#0a0a0a' : '#888', width: 26, height: 28, borderRadius: 4, fontSize: 11, fontWeight: 700 }}>{n}</button>
      ))}
    </div>
  );
}

// ============================================================
// HISTORY VIEW
// ============================================================

function HistoryView({ history, photos, onZoom }) {
  const sortedDates = Object.keys(history).sort().reverse();
  const totalSessions = sortedDates.length;
  const totalSets = sortedDates.reduce((sum, d) => sum + Object.values(history[d].sets || {}).reduce((s, sets) => s + sets.length, 0), 0);
  if (sortedDates.length === 0) return (
    <div style={{ textAlign: 'center', padding: 36, color: '#666' }}>
      <Calendar size={32} style={{ marginBottom: 10, opacity: 0.5 }} />
      <div style={{ fontSize: 13 }}>No workouts logged yet</div>
      <div style={{ fontSize: 11, marginTop: 6 }}>Logged sets will appear here. If you have data from a previous backup, paste it via the RESTORE button at the top.</div>
    </div>
  );
  return (
    <div>
      <div className="display" style={{ fontSize: 26, color: '#facc15', marginBottom: 4 }}>HISTORY</div>
      <div style={{ fontSize: 11, color: '#888', marginBottom: 12 }}>{totalSessions} sessions · {totalSets} sets logged</div>
      {sortedDates.map(key => {
        const day = history[key];
        const date = day.date || key.split('__')[0];
        const dayMeta = PROGRAM[day.day];
        const dayLabel = dayMeta ? `${dayMeta.label}: ${dayMeta.name}` : (day.day || '').toString().toUpperCase();
        const totalSets = Object.values(day.sets || {}).reduce((sum, sets) => sum + sets.length, 0);
        const duration = day.endTime && day.startTime ? Math.round((new Date(day.endTime) - new Date(day.startTime)) / 60000) : null;
        return (
          <div key={key} style={{ background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 11, marginBottom: 9 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 7 }}>
              <div>
                <div className="mono" style={{ fontSize: 9.5, color: '#666' }}>{date}</div>
                <div className="display" style={{ fontSize: 17, color: '#facc15', lineHeight: 1.1 }}>{dayLabel}</div>
              </div>
              <div style={{ textAlign: 'right' }}>
                <div className="mono" style={{ fontSize: 11, color: '#ddd' }}>{totalSets} SETS</div>
                {duration && <div className="mono" style={{ fontSize: 9.5, color: '#666' }}>{duration} MIN</div>}
              </div>
            </div>
            <div>
              {Object.entries(day.sets || {}).map(([exId, sets]) => {
                const ex = Object.values(PROGRAM).flatMap(d => d.exercises || []).find(e => e.id === exId);
                const photo = photos[exId];
                return (
                  <div key={exId} style={{ marginBottom: 5, fontSize: 10.5, display: 'flex', gap: 8, alignItems: 'flex-start' }}>
                    {photo && <div onClick={() => onZoom?.(photo)} style={{ width: 24, height: 24, borderRadius: 4, background: `url(${photo}) center/cover`, flexShrink: 0, cursor: 'zoom-in' }} />}
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ color: '#aaa', fontWeight: 700 }}>{ex?.name || exId}</div>
                      <div className="mono" style={{ color: '#666', fontSize: 9.5, marginLeft: 2 }}>
                        {sets.map(s => `${s.weight}${s.weightMode === 'perSide' ? '×2' : ''}${s.unit}×${s.reps}`).join('  ·  ')}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// ============================================================
// PROGRESS VIEW
// ============================================================

function ProgressView({ history, bodyweightLog, setBodyweightLog, settings }) {
  const [bwInput, setBwInput] = useState('');
  const allExercises = Object.values(PROGRAM).flatMap(d => d.exercises || []);
  const exerciseProgress = allExercises.map(ex => {
    const sessions = Object.entries(history)
      .filter(([, day]) => day.sets?.[ex.id]?.length)
      .map(([date, day]) => {
        const sets = day.sets[ex.id];
        const top = sets.reduce((a, b) => (a.weight * a.reps > b.weight * b.reps ? a : b));
        return { date, topVolume: top.weight * top.reps, weight: top.weight, reps: top.reps, unit: top.unit };
      })
      .sort((a, b) => a.date.localeCompare(b.date));
    return { ex, sessions };
  }).filter(p => p.sessions.length > 0);

  const logBw = () => {
    if (!bwInput) return;
    setBodyweightLog([...bodyweightLog, { date: todayKey(), weight: parseFloat(bwInput), unit: settings.unit }]);
    setBwInput('');
  };

  return (
    <div>
      <div className="display" style={{ fontSize: 26, color: '#facc15', marginBottom: 12 }}>PROGRESS</div>
      <div style={{ background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 11, marginBottom: 12 }}>
        <div style={{ fontSize: 10.5, color: '#facc15', fontWeight: 700, letterSpacing: '0.05em', marginBottom: 9 }}>BODYWEIGHT</div>
        <div style={{ display: 'flex', gap: 7, marginBottom: 8 }}>
          <input type="number" inputMode="decimal" value={bwInput} onChange={e => setBwInput(e.target.value)} placeholder={`Today (${settings.unit})`} style={{ flex: 1, background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 9, borderRadius: 7, fontSize: 13, fontFamily: 'JetBrains Mono, monospace' }} />
          <button onClick={logBw} style={{ background: '#facc15', color: '#0a0a0a', padding: '9px 14px', borderRadius: 7, fontWeight: 700, fontSize: 11 }}>LOG</button>
        </div>
        {bodyweightLog.length > 0 && (
          <div style={{ maxHeight: 110, overflowY: 'auto' }}>
            {bodyweightLog.slice().reverse().slice(0, 5).map((b, i) => (
              <div key={i} className="mono" style={{ fontSize: 10.5, color: '#888', padding: '3px 0', borderBottom: '1px solid #1a1a1a' }}>{b.date} · {b.weight} {b.unit}</div>
            ))}
          </div>
        )}
      </div>
      {exerciseProgress.length === 0 ? (
        <div style={{ textAlign: 'center', padding: 26, color: '#666', fontSize: 11 }}>Log workouts to see progression</div>
      ) : exerciseProgress.map(({ ex, sessions }) => {
        const first = sessions[0], last = sessions[sessions.length - 1];
        const change = ((last.topVolume - first.topVolume) / first.topVolume * 100).toFixed(1);
        const positive = parseFloat(change) >= 0;
        return (
          <div key={ex.id} style={{ background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 11, marginBottom: 9 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 5 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: '#f5f5f5' }}>{ex.name}</div>
              {sessions.length > 1 && <span className="mono" style={{ fontSize: 10.5, color: positive ? '#22c55e' : '#ef4444', fontWeight: 700 }}>{positive ? '+' : ''}{change}%</span>}
            </div>
            <Sparkline data={sessions.map(s => s.topVolume)} />
            <div className="mono" style={{ fontSize: 9.5, color: '#666', marginTop: 4 }}>{sessions.length} sessions · Latest: {last.weight}{settings.unit} × {last.reps}</div>
          </div>
        );
      })}
    </div>
  );
}

function Sparkline({ data }) {
  if (data.length < 2) return <div style={{ fontSize: 10, color: '#666' }}>Need 2+ sessions</div>;
  const max = Math.max(...data), min = Math.min(...data), range = max - min || 1;
  const W = 280, H = 40;
  const step = W / (data.length - 1);
  const points = data.map((v, i) => `${i * step},${H - ((v - min) / range) * H}`).join(' ');
  return (
    <svg viewBox={`0 0 ${W} ${H}`} style={{ width: '100%', height: 40 }}>
      <polyline points={points} fill="none" stroke="#facc15" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
      {data.map((v, i) => <circle key={i} cx={i * step} cy={H - ((v - min) / range) * H} r="2.5" fill="#facc15" />)}
    </svg>
  );
}

// ============================================================
// SETTINGS VIEW
// ============================================================

function SyncSettings({ status, lastSyncAt, onSave, onSyncNow }) {
  const [open, setOpen] = useState(!sync.isConfigured());
  const [token, setToken] = useState(sync.getToken());
  const [repo, setRepo] = useState(sync.getRepo());
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState(null);
  const save = async () => { setBusy(true); setMsg(null); const r = await onSave(token, repo); setMsg(r); setBusy(false); if (r.ok) { setToken(''); setOpen(false); } };
  const color = status === 'synced' ? '#22c55e' : status === 'syncing' ? '#facc15' : status === 'off' ? '#666' : status === 'offline' ? '#888' : '#ef4444';
  const label = status === 'synced' ? 'Synced' : status === 'syncing' ? 'Syncing…' : status === 'off' ? 'Not connected' : status === 'offline' ? 'Offline — saved on this device' : status === 'error' ? 'Error' : 'Idle';
  return (
    <div style={{ marginBottom: 10, background: '#111', border: '2px solid #2f6fed', borderRadius: 10, padding: 12 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ fontSize: 10.5, color: '#60a5fa', fontWeight: 700, letterSpacing: '0.05em', display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ width: 8, height: 8, borderRadius: 4, background: color }} /> ☁ GITHUB SYNC — PHONE + MAC
        </div>
        <button onClick={() => setOpen(o => !o)} style={{ background: 'none', color: '#888', fontSize: 10, fontWeight: 700 }}>{open ? 'HIDE' : 'SETUP'}</button>
      </div>
      <div style={{ fontSize: 10, color: '#888', marginTop: 4 }}>{label}{lastSyncAt ? ` · last ${new Date(lastSyncAt).toLocaleTimeString()}` : ''}</div>
      {sync.isConfigured() && <button onClick={onSyncNow} style={{ marginTop: 8, width: '100%', background: '#1a1a1a', color: '#60a5fa', border: '1px solid #2f6fed', padding: 8, borderRadius: 7, fontWeight: 700, fontSize: 11 }}>↻ SYNC NOW</button>}
      {open && (
        <div style={{ marginTop: 10 }}>
          <div style={{ fontSize: 9.5, color: '#888', lineHeight: 1.55, marginBottom: 8 }}>
            Make a GitHub <b style={{ color: '#ddd' }}>fine-grained token</b> with <b style={{ color: '#ddd' }}>Contents: Read &amp; Write</b> on the <b style={{ color: '#ddd' }}>hypertrophy-data</b> repo only, then paste it below. It's stored only in this browser, never uploaded. <a href="https://github.com/settings/personal-access-tokens/new" target="_blank" rel="noopener noreferrer" style={{ color: '#60a5fa' }}>Create token →</a>
          </div>
          <input type="text" value={repo} onChange={e => setRepo(e.target.value)} placeholder="owner/repo" style={{ width: '100%', background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 8, borderRadius: 6, fontSize: 11, fontFamily: 'JetBrains Mono, monospace', marginBottom: 6 }} />
          <input type="password" value={token} onChange={e => setToken(e.target.value)} placeholder="github_pat_…" autoComplete="off" style={{ width: '100%', background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 8, borderRadius: 6, fontSize: 11, fontFamily: 'JetBrains Mono, monospace', marginBottom: 6 }} />
          <button onClick={save} disabled={busy} style={{ width: '100%', background: '#2f6fed', color: '#fff', padding: 9, borderRadius: 7, fontWeight: 700, fontSize: 11 }}>{busy ? 'CONNECTING…' : 'CONNECT & SYNC'}</button>
          {msg && <div style={{ marginTop: 7, fontSize: 10, color: msg.ok ? '#22c55e' : '#fca5a5' }}>{msg.msg}</div>}
        </div>
      )}
    </div>
  );
}

function SettingsView({ settings, setSettings, onExportExcel, onExportJSON, onImportJSON, onImportExcel, onBackupNow, onRestore, lastBackup, syncStatus, lastSyncAt, onSaveSync, onSyncNow }) {
  const importRef = useRef(null);
  const excelImportRef = useRef(null);
  return (
    <div>
      <div className="display" style={{ fontSize: 26, color: '#facc15', marginBottom: 12 }}>SETTINGS</div>
      <SyncSettings status={syncStatus} lastSyncAt={lastSyncAt} onSave={onSaveSync} onSyncNow={onSyncNow} />
      <SettingRow label="UNIT">
        <div style={{ display: 'flex', background: '#0a0a0a', borderRadius: 7, padding: 2, border: '1px solid #2a2a2a' }}>
          {['kg', 'lb'].map(u => (
            <button key={u} onClick={() => setSettings({ ...settings, unit: u })} style={{ padding: '7px 14px', background: settings.unit === u ? '#facc15' : 'transparent', color: settings.unit === u ? '#0a0a0a' : '#888', borderRadius: 5, fontWeight: 700, fontSize: 11 }}>{u.toUpperCase()}</button>
          ))}
        </div>
      </SettingRow>
      <SettingRow label="BODYWEIGHT">
        <input type="number" inputMode="decimal" value={settings.bodyweight} onChange={e => setSettings({ ...settings, bodyweight: parseFloat(e.target.value) || 0 })} style={{ width: 70, background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 7, borderRadius: 5, fontSize: 12, fontFamily: 'JetBrains Mono, monospace', textAlign: 'center' }} />
        <span style={{ marginLeft: 7, color: '#888', fontSize: 11 }}>{settings.unit}</span>
      </SettingRow>
      <SettingRow label="AGE (HR zones)">
        <input type="number" inputMode="numeric" value={settings.age} onChange={e => setSettings({ ...settings, age: parseInt(e.target.value) || 30 })} style={{ width: 55, background: '#0a0a0a', border: '1px solid #2a2a2a', color: '#f5f5f5', padding: 7, borderRadius: 5, fontSize: 12, fontFamily: 'JetBrains Mono, monospace', textAlign: 'center' }} />
      </SettingRow>
      <SettingRow label="MESOCYCLE">
        <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
          <button onClick={() => setSettings({ ...settings, mesoWeek: Math.max(1, settings.mesoWeek - 1) })} style={{ background: '#1a1a1a', color: '#aaa', width: 26, height: 26, borderRadius: 5 }}><Minus size={11} /></button>
          <span className="mono" style={{ fontSize: 15, fontWeight: 700, color: '#facc15', width: 20, textAlign: 'center' }}>{settings.mesoWeek}</span>
          <button onClick={() => setSettings({ ...settings, mesoWeek: Math.min(6, settings.mesoWeek + 1) })} style={{ background: '#1a1a1a', color: '#aaa', width: 26, height: 26, borderRadius: 5 }}><Plus size={11} /></button>
          {settings.mesoWeek === 6 && <span style={{ color: '#ef4444', fontSize: 9.5, fontWeight: 700, marginLeft: 3 }}>DELOAD</span>}
        </div>
      </SettingRow>

      {/* CLIPBOARD BACKUP — primary */}
      <div style={{ marginTop: 16, background: '#111', border: '2px solid #facc15', borderRadius: 10, padding: 12 }}>
        <div style={{ fontSize: 10.5, color: '#facc15', fontWeight: 700, letterSpacing: '0.05em', marginBottom: 3 }}>📋 CLIPBOARD BACKUP (RECOMMENDED)</div>
        <div style={{ fontSize: 10, color: '#888', marginBottom: 10, lineHeight: 1.4 }}>Copies all your data to clipboard. Paste it into Notes, email to yourself, or save to Google Drive. Restore by copying the text and tapping RESTORE.</div>
        <button onClick={onBackupNow} style={{ width: '100%', background: '#facc15', color: '#0a0a0a', padding: 10, borderRadius: 7, fontWeight: 700, fontSize: 12, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, marginBottom: 7 }}><Copy size={12} /> BACKUP TO CLIPBOARD</button>
        <button onClick={onRestore} style={{ width: '100%', background: '#1a1a1a', color: '#facc15', border: '1px solid #facc15', padding: 10, borderRadius: 7, fontWeight: 700, fontSize: 12, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6 }}><ClipboardPaste size={12} /> RESTORE FROM CLIPBOARD</button>
        {lastBackup && <div style={{ marginTop: 8, fontSize: 10, color: '#22c55e', textAlign: 'center' }}>Last backed up: {new Date(lastBackup).toLocaleString()}</div>}
      </div>

      {/* EXCEL */}
      <div style={{ marginTop: 10, background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 12 }}>
        <div style={{ fontSize: 10.5, color: '#facc15', fontWeight: 700, letterSpacing: '0.05em', marginBottom: 3 }}>EXCEL — EDITABLE</div>
        <div style={{ fontSize: 10, color: '#888', marginBottom: 10, lineHeight: 1.4 }}>Export to spreadsheet, edit weights or add Garmin data, re-import.</div>
        <button onClick={onExportExcel} style={{ width: '100%', background: '#1a1a1a', color: '#aaa', padding: 9, borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, marginBottom: 6, border: '1px solid #2a2a2a' }}><Download size={11} /> EXPORT EXCEL</button>
        <button onClick={() => excelImportRef.current?.click()} style={{ width: '100%', background: '#1a1a1a', color: '#aaa', padding: 9, borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, border: '1px solid #2a2a2a' }}><Upload size={11} /> IMPORT EXCEL</button>
        <input ref={excelImportRef} type="file" accept=".xlsx,.xls" onChange={e => { if (e.target.files?.[0]) { onImportExcel(e.target.files[0]); e.target.value = ''; } }} style={{ display: 'none' }} />
      </div>

      {/* JSON FILE */}
      <div style={{ marginTop: 10, background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 12 }}>
        <div style={{ fontSize: 10.5, color: '#888', fontWeight: 700, letterSpacing: '0.05em', marginBottom: 3 }}>JSON FILE BACKUP</div>
        <div style={{ fontSize: 10, color: '#666', marginBottom: 10, lineHeight: 1.4 }}>Same as clipboard but downloads as a file. Use this if your phone allows file downloads.</div>
        <button onClick={onExportJSON} style={{ width: '100%', background: '#1a1a1a', color: '#aaa', padding: 9, borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, border: '1px solid #2a2a2a', marginBottom: 6 }}><Download size={11} /> DOWNLOAD JSON</button>
        <button onClick={() => importRef.current?.click()} style={{ width: '100%', background: '#1a1a1a', color: '#aaa', padding: 9, borderRadius: 7, fontWeight: 700, fontSize: 11, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5, border: '1px solid #2a2a2a' }}><Upload size={11} /> RESTORE FROM FILE</button>
        <input ref={importRef} type="file" accept=".json" onChange={e => { if (e.target.files?.[0]) { onImportJSON(e.target.files[0]); e.target.value = ''; } }} style={{ display: 'none' }} />
      </div>

      <div style={{ marginTop: 10, background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 12 }}>
        <div style={{ fontSize: 10.5, color: '#facc15', fontWeight: 700, letterSpacing: '0.05em', marginBottom: 7 }}>GARMIN FENIX 8</div>
        <div style={{ fontSize: 11, color: '#ccc', lineHeight: 1.6 }}>
          <div>• Start <strong>Strength Training</strong> activity at warmup</div>
          <div>• Zone 2 ceiling: <strong className="mono" style={{ color: '#facc15' }}>{Math.round((220 - settings.age) * 0.7)} bpm</strong></div>
          <div>• Body Battery &lt;30 = deload signal</div>
          <div>• HRV down 5+ days = overreaching</div>
        </div>
        <div style={{ marginTop: 10, padding: 8, background: '#0a0a0a', borderRadius: 6, fontSize: 10, color: '#ccc', lineHeight: 1.5 }}>
          <strong style={{ color: '#facc15' }}>To import from Garmin:</strong> Export Excel from this app → open in Sheets → fill in past weights from Garmin → re-import via "IMPORT EXCEL" above.
        </div>
      </div>

      <div style={{ marginTop: 14, fontSize: 9.5, color: '#444', textAlign: 'center' }}>v5.1 · Upper-Focus · Push / Pull / Lower / Upper</div>
    </div>
  );
}

function SettingRow({ label, children }) {
  return (
    <div style={{ background: '#111', border: '1px solid #1f1f1f', borderRadius: 10, padding: 11, marginBottom: 7, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
      <div style={{ fontSize: 10, color: '#888', fontWeight: 700, letterSpacing: '0.05em' }}>{label}</div>
      <div style={{ display: 'flex', alignItems: 'center' }}>{children}</div>
    </div>
  );
}
