// Combat Extender for Wrath & Glory
// 
// VERSION: v1.2 - Aim & Engagement Bug Fixes
//
// CHANGES FROM v1.1:
// - Fixed duplicate Aim bonus (was adding +2 instead of +1)
// - Fixed short range suppression on first open (use dialog.fields.range instead of DOM cache)  
// - Trigger recompute after first patch to fix first-open timing issue
//
// PREVIOUS: v1.1 - Fixed infinite loop bug with render guard

import {
  COMBAT_OPTION_LABELS,
  COVER_DIFFICULTY_VALUES,
  ENGAGED_TOOLTIP_LABELS,
  MODULE_BASE_PATH,
  MODULE_ID,
  TEMPLATE_BASE_PATH,
  VISION_PENALTIES,
  SIZE_MODIFIER_OPTIONS,
  SIZE_OPTION_KEYS
} from "./constants.js";
import { getEngagedEffect, isActiveScene } from "./engagement.js";
import { log, logDebug, logError } from "./logging.js";
import {
  getCanvasMeasurementContext,
  getCoverDifficulty,
  getCoverLabel,
  normalizeCoverKey,
  normalizeSizeKey,
  measureTokenDistance,
  tokensAreEngaged,
  tokensAreEngagedUsingDistance
} from "./measurement.js";
import { syncAllOutAttackCondition } from "./turn-effects.js";

function getTargetSize(dialog) {
  const target = dialog?.data?.targets?.[0];
  const actor = target?.actor ?? target?.document?.actor;
  if (!actor) return "average";

  const size = actor.system?.combat?.size ?? actor.system?.size ?? actor.size;
  return normalizeSizeKey(size);
}

function getTargetResolve(dialog) {
  const target = dialog?.data?.targets?.[0];
  const actor = target?.actor ?? target?.document?.actor;
  if (!actor) return null;

  const resolveTotal = Number(foundry.utils.getProperty(actor.system, "attributes.resolve.total"));
  if (Number.isFinite(resolveTotal) && resolveTotal > 0) return resolveTotal;

  const resolveValue = Number(foundry.utils.getProperty(actor.system, "attributes.resolve.value"));
  if (Number.isFinite(resolveValue) && resolveValue > 0) return resolveValue;

  return null;
}

function getTargetIdentifier(dialog) {
  const target = dialog?.data?.targets?.[0];
  if (!target) return null;
  return target?.document?.id ?? target?.id ?? target?.token?.id ?? null;
}

function resolvePlaceableToken(tokenLike, { requireActiveScene = false } = {}) {
  if (!tokenLike) return null;

  let token = null;
  if (tokenLike.center && tokenLike.document) {
    token = tokenLike;
  } else if (tokenLike.object?.center && tokenLike.object?.document) {
    token = tokenLike.object;
  } else if (tokenLike.document?.object?.center && tokenLike.document?.object?.document) {
    token = tokenLike.document.object;
  } else if (tokenLike.token) {
    token = resolvePlaceableToken(tokenLike.token, { requireActiveScene: false });
  }

  if (!token) return null;

  if (requireActiveScene) {
    const sceneRef = token.scene ?? token.document?.parent ?? token.parent ?? null;
    if (sceneRef && !isActiveScene(sceneRef)) return null;
  }

  return token;
}

function getDialogAttackerToken(dialog) {
  if (!dialog) return null;

  const directToken = resolvePlaceableToken(dialog.token, { requireActiveScene: true });
  if (directToken) return directToken;

  const actor = dialog.actor ?? dialog.token?.actor ?? null;
  if (!actor) return null;

  const activeTokens = typeof actor.getActiveTokens === "function"
    ? actor.getActiveTokens(true)
    : [];

  for (const candidate of activeTokens) {
    const resolved = resolvePlaceableToken(candidate, { requireActiveScene: true });
    if (resolved) return resolved;
  }

  return null;
}

function getDialogTargetTokens(dialog) {
  const targets = Array.isArray(dialog?.data?.targets) ? dialog.data.targets : [];
  if (!targets.length) return [];

  const results = [];
  const seen = new Set();

  for (const entry of targets) {
    const token = resolvePlaceableToken(entry, { requireActiveScene: true });
    if (!token) continue;

    const identifier = token.id ?? token.document?.id ?? token.document?.uuid ?? null;
    if (identifier) {
      if (seen.has(identifier)) continue;
      seen.add(identifier);
    }

    results.push(token);
  }

  return results;
}

function combatOptionsActive(fields) {
  if (!fields) return false;

  return Boolean(
    fields.allOutAttack ||
    fields.charging ||
    fields.aim ||
    fields.brace ||
    fields.pinning ||
    normalizeCoverKey(fields.cover ?? "") ||
    fields.pistolsInMelee ||
    normalizeSizeKey(fields.sizeModifier ?? "") ||
    fields.visionPenalty ||
    fields.disarm ||
    fields.calledShot?.enabled ||
    normalizeSizeKey(fields.calledShot?.size ?? "")
  );
}

Hooks.once("init", async () => {
  await loadTemplates([
    `${TEMPLATE_BASE_PATH}/combat-options.hbs`,
    `${TEMPLATE_BASE_PATH}/partials/co-checkbox.hbs`,
    `${TEMPLATE_BASE_PATH}/partials/co-select.hbs`
  ]);

  Handlebars.registerPartial("co-checkbox", await fetch(`${TEMPLATE_BASE_PATH}/partials/co-checkbox.hbs`).then(r => r.text()));
  Handlebars.registerPartial("co-select", await fetch(`${TEMPLATE_BASE_PATH}/partials/co-select.hbs`).then(r => r.text()));

  Handlebars.registerHelper("t", (s) => String(s));
  Handlebars.registerHelper("eq", (a, b) => a === b);
  Handlebars.registerHelper("not", (v) => !v);
  Handlebars.registerHelper("concat", (...a) => a.slice(0, -1).join(""));
});

const patchedWeaponDialogPrototypes = new WeakSet();

function ensureWeaponDialogPatched(app) {
  const prototype = app?.constructor?.prototype ?? Object.getPrototypeOf(app);
  if (!prototype || prototype === Application.prototype) return false;
  if (patchedWeaponDialogPrototypes.has(prototype)) return false;

  const originalPrepareContext = prototype._prepareContext;
  const originalDefaultFields  = prototype._defaultFields;
  const originalComputeFields = prototype.computeFields;
  const originalGetSubmissionData = prototype._getSubmissionData;

  if (typeof originalPrepareContext !== "function" ||
      typeof originalDefaultFields  !== "function" ||
      typeof originalGetSubmissionData !== "function" ||
      typeof originalComputeFields !== "function") {
    logError("WeaponDialog prototype missing expected methods");
    return false;
  }

  prototype._prepareContext = async function (options) {
    const context = await originalPrepareContext.call(this, options);

    context.coverOptions = {
      "": "No Cover",
      half: COMBAT_OPTION_LABELS.halfCover,
      full: COMBAT_OPTION_LABELS.fullCover
    };

    const weapon = this.weapon;
    const salvoValue = Number(weapon?.system?.salvo ?? weapon?.salvo ?? 0);
    const canPinning = Boolean(weapon?.isRanged) && Number.isFinite(salvoValue) && salvoValue > 1;

    const fields = this.fields ?? (this.fields = {});
    if (!canPinning && fields.pinning) {
      fields.pinning = false;
    }

    const actor = this.actor ?? this.token?.actor ?? null;
    const isEngaged = Boolean(getEngagedEffect(actor));
    const pistolTrait = weapon?.system?.traits;
    const hasPistolTrait = Boolean(pistolTrait?.has?.("pistol") || pistolTrait?.get?.("pistol"));
    const canPistolsInMelee = hasPistolTrait && isEngaged;
    this._combatOptionsCanPistolsInMelee = canPistolsInMelee;

    return context;
  };

  prototype._defaultFields = function () {
    const baseFields = originalDefaultFields.call(this);
    return foundry.utils.mergeObject(baseFields, {
      cover: "",
      visionPenalty: "",
      sizeModifier: "",
      allOutAttack: false,
      charging: false,
      aim: false,
      brace: false,
      pinning: false,
      pistolsInMelee: false,
      disarm: false,
      calledShot: {
        enabled: false,
        size: ""
      }
    });
  };

  prototype._getSubmissionData = function () {
    const submitData = foundry.utils.mergeObject(this.data ?? {}, this.fields ?? {});
    submitData.context = this.context;

    if (!this.context?.skipTargets) {
      const rawTargets = Array.isArray(submitData.targets)
        ? submitData.targets
        : Array.from(submitData.targets ?? []);

      submitData.targets = rawTargets
        .map((target) => {
          const actor = resolveTargetActor(target);
          if (!actor) return null;

          const token = resolveTargetToken(target, actor);
          const tokenDocument = token?.document ?? token ?? null;

          return typeof actor.speakerData === "function"
            ? actor.speakerData(tokenDocument)
            : null;
        })
        .filter(Boolean);
    }

    if (typeof this.createBreakdown === "function") {
      submitData.context.breakdown = this.createBreakdown();
    }

    return submitData;
  };

  prototype.computeFields = async function (...args) {
    const result = await originalComputeFields.apply(this, args);

    this._combatExtenderSystemBaseline = foundry.utils.deepClone(this.fields ?? {});

    try {
      await applyCombatExtender(this);
    } catch (err) {
      logError("Combat Extender computeFields patch failed", err);
    }

    return result ?? this.fields;
  };

  patchedWeaponDialogPrototypes.add(prototype);
  return true;
}

function resolveTargetActor(target) {
  if (!target) return null;

  const targetActor = target.actor ?? target.document?.actor;
  if (targetActor) return targetActor;

  const actorId = target.actorId ?? target.actor?.id;
  if (actorId) {
    const actor = game.actors?.get?.(actorId);
    if (actor) return actor;
  }

  return resolveTargetToken(target)?.actor ?? null;
}

function resolveTargetToken(target, actor) {
  if (!target) return null;

  const token = target.document ?? target.token;
  if (token) return token;

  const tokenId = target.documentId ?? target.tokenId ?? target.id;
  const sceneId = target.sceneId ?? target.scene?.id ?? target.document?.parent?.id;
  if (tokenId && sceneId) {
    const scene = game.scenes?.get?.(sceneId);
    const foundToken = scene?.tokens?.get?.(tokenId);
    if (foundToken) return foundToken;
  }

  const actorToken = actor?.getActiveTokens?.()?.at?.(0);
  if (actorToken) return actorToken.document ?? actorToken;

  return null;
}

function syncDialogInputsFromFields(app, html) {
  const fields = app?.fields ?? {};
  const $html = html instanceof jQuery ? html : $(html);

  $html.find("input[name], select[name]").each((_, el) => {
    const name = el.name;
    if (!name) return;

    const value = foundry.utils.getProperty(fields, name);
    if (value === undefined) return;

    if (el.type === "checkbox") {
      el.checked = Boolean(value);
      return;
    }

    const stringValue = value == null ? "" : String(value);
    if (el.value !== stringValue) {
      el.value = stringValue;
    }
  });
}

async function applyCombatExtender(dialog) {
  const weapon = dialog.weapon;
  if (!weapon) return;

  console.log("=== applyCombatExtender START ===");

  const fields = dialog.fields ?? (dialog.fields = {});

  const systemBaseline = dialog._combatExtenderSystemBaseline;
  if (systemBaseline) {
    const safeBaseline = foundry.utils.deepClone(systemBaseline);
    if (safeBaseline.pool !== undefined) fields.pool = safeBaseline.pool;
    if (safeBaseline.difficulty !== undefined) fields.difficulty = safeBaseline.difficulty;
    if (safeBaseline.damage !== undefined) fields.damage = safeBaseline.damage;
    if (safeBaseline.ed !== undefined) fields.ed = foundry.utils.deepClone(safeBaseline.ed);
    if (safeBaseline.ap !== undefined) fields.ap = foundry.utils.deepClone(safeBaseline.ap);
    if (safeBaseline.wrath !== undefined) fields.wrath = safeBaseline.wrath;
  }

  fields.ed = foundry.utils.mergeObject({ value: 0, dice: 0 }, fields.ed ?? {}, { inplace: false });
  fields.ap = foundry.utils.mergeObject({ value: 0, dice: 0 }, fields.ap ?? {}, { inplace: false });
  if (fields.damage === undefined) fields.damage = 0;

  logDebug("CE fields at start:", { 
    cover: fields.cover, 
    visionPenalty: fields.visionPenalty, 
    sizeModifier: fields.sizeModifier, 
    pool: fields.pool, 
    difficulty: fields.difficulty 
  });

  const systemBaselineSnapshot = foundry.utils.deepClone(fields ?? {});
  const systemDifficulty = Number.isFinite(systemBaselineSnapshot.difficulty)
    ? Number(systemBaselineSnapshot.difficulty)
    : null;

  const manualOverridesRaw = dialog._combatOptionsManualOverrides
    ? foundry.utils.deepClone(dialog._combatOptionsManualOverrides)
    : null;
  const manualOverrides = manualOverridesRaw && Object.keys(manualOverridesRaw).length
    ? manualOverridesRaw
    : null;
  
  let pool = Number(fields.pool ?? 0);
  let difficulty = Number(fields.difficulty ?? 0);
  let damage = fields.damage ?? 0;
  let edValue = Number(fields.ed?.value ?? 0);
  let edDice = Number(fields.ed?.dice ?? 0);
  let apValue = Number(fields.ap?.value ?? 0);
  let apDice = Number(fields.ap?.dice ?? 0);
  let wrath = Number(fields.wrath ?? 0);

  console.log("CE: Initial values - pool:", pool, "difficulty:", difficulty, "damage:", damage, "aim:", fields.aim);

  const baseDamage = damage;
  const baseEdValue = edValue;
  const baseEdDice = edDice;

  const addTooltip = (field, value, label) => {
    // Tooltip implementation
  };

  // --- Pistols while Engaged ---
  const actor = dialog.actor ?? dialog.token?.actor ?? null;
  const isEngaged = Boolean(getEngagedEffect(actor));

  const traits = weapon?.system?.traits;
  const hasPistol = Boolean(traits?.has?.("pistol") || traits?.get?.("pistol"));

  // Use dialog.fields.range - it's calculated by system before computeFields runs
  // DOM cache (_combatExtenderRangeBand) isn't set yet on first dialog open
  const rangeBand = String(dialog.fields?.range ?? "").toLowerCase();

  console.log("CE: Engagement check - isEngaged:", isEngaged, "hasPistol:", hasPistol, "rangeBand:", rangeBand);

  if (isEngaged && weapon?.isRanged && hasPistol) {
    console.log("CE: Applying engagement penalties");
    // +2 DN when firing pistols while engaged
    difficulty += 2;
    addTooltip("difficulty", 2, "Engaged + Pistol (+2 DN)");

    // Cannot Aim while engaged
    if (fields.aim) fields.aim = false;

    // Short range bonus die is not allowed while engaged
    if (rangeBand === "short") {
      pool -= 1;
      addTooltip("pool", -1, "Short Range suppressed (Engaged + Pistol)");
    }

    fields.pistolsInMelee = true;
  }
  // --- end pistols while engaged ---

  let damageSuppressed = false;
  let restoreTargetSizeTooltip = null;

  if (fields.allOutAttack) {
    pool += 2;
    addTooltip("pool", 2, COMBAT_OPTION_LABELS.allOutAttack);
    logDebug("CE all-out attack:", { previousPool: pool - 2, nextPool: pool });
  }

  if (fields.charging) {
    pool += 1;
    addTooltip("pool", 1, COMBAT_OPTION_LABELS.charge);
    logDebug("CE charge:", { previousPool: pool - 1, nextPool: pool });
  }

  // NOTE: Aim bonus (+1) is already added by the system's computeFields()
  // Combat Extender should only SUPPRESS it when engaged, not add it again
  // Removed duplicate: if (fields.aim) { pool += 1; }

  const visionKey = fields.visionPenalty;
  const visionPenaltyData = VISION_PENALTIES[visionKey];
  if (visionPenaltyData) {
    const previousDifficulty = difficulty;
    const penalty = weapon?.isMelee ? visionPenaltyData.melee : visionPenaltyData.ranged;
    if (penalty > 0) difficulty += penalty;
    addTooltip("difficulty", penalty ?? 0, visionPenaltyData.label);
    logDebug("CE vision modifier:", { visionKey, penalty, previousDifficulty, nextDifficulty: difficulty });
  }

  const sizeKey = fields.sizeModifier;
  const sizeModifierData = SIZE_MODIFIER_OPTIONS[sizeKey];
  if (sizeModifierData) {
    const previousPool = pool;
    const previousDifficulty = difficulty;
    if (sizeModifierData.pool) {
      pool += sizeModifierData.pool;
      addTooltip("pool", sizeModifierData.pool, sizeModifierData.label);
    }
    if (sizeModifierData.difficulty) {
      difficulty += sizeModifierData.difficulty;
      addTooltip("difficulty", sizeModifierData.difficulty, sizeModifierData.label);
    }
    logDebug("CE size modifier:", { sizeKey, previousPool, nextPool: pool, previousDifficulty, nextDifficulty: difficulty });
  }

  if (fields.disarm) {
    if (baseDamage) addTooltip("damage", -baseDamage, COMBAT_OPTION_LABELS.calledShotDisarm);
    addTooltip("difficulty", 0, COMBAT_OPTION_LABELS.calledShotDisarm);
    damage = 0;
    edValue = 0;
    edDice = 0;
    damageSuppressed = true;
  }

  const statusCover = normalizeCoverKey(dialog._combatOptionsDefaultCover ?? "");
  const selectedCover = normalizeCoverKey(fields.cover);
  const coverDelta = getCoverDifficulty(selectedCover) - getCoverDifficulty(statusCover);

  if (coverDelta !== 0) {
    difficulty += coverDelta;
    const label = getCoverLabel(coverDelta > 0 ? selectedCover : statusCover);
    if (label) addTooltip("difficulty", coverDelta, label);
  }

  logDebug("CE cover modifier:", { statusCover, selectedCover, coverDelta, nextDifficulty: difficulty });

  if (!damageSuppressed) {
    damage = baseDamage;
    edValue = baseEdValue;
    edDice = baseEdDice;
  }

  const delta = {
    pool: pool - Number(systemBaselineSnapshot.pool ?? 0),
    difficulty: difficulty - Number(systemBaselineSnapshot.difficulty ?? 0),
    damage: damage - (systemBaselineSnapshot.damage ?? 0),
    ed: {
      value: edValue - Number(systemBaselineSnapshot.ed?.value ?? 0),
      dice: edDice - Number(systemBaselineSnapshot.ed?.dice ?? 0)
    },
    ap: {
      value: apValue - Number(systemBaselineSnapshot.ap?.value ?? 0),
      dice: apDice - Number(systemBaselineSnapshot.ap?.dice ?? 0)
    }
  };
  dialog._combatExtenderDelta = delta;

  const finalPool = manualOverrides?.pool !== undefined
    ? Math.max(0, Number(manualOverrides.pool ?? 0))
    : Math.max(0, pool);
  const finalDifficulty = manualOverrides?.difficulty !== undefined
    ? Math.max(0, Number(manualOverrides.difficulty ?? 0))
    : Math.max(0, difficulty);
  const finalEd = manualOverrides?.ed !== undefined
    ? foundry.utils.deepClone(manualOverrides.ed)
    : { value: edValue, dice: edDice };
  finalEd.value = Math.max(0, Number(finalEd?.value ?? 0));
  finalEd.dice = Math.max(0, Number(finalEd?.dice ?? 0));

  const finalAp = manualOverrides?.ap !== undefined
    ? foundry.utils.deepClone(manualOverrides.ap)
    : { value: apValue, dice: apDice };
  finalAp.value = Math.max(0, Number(finalAp?.value ?? 0));
  finalAp.dice = Math.max(0, Number(finalAp?.dice ?? 0));

  const finalWrath = manualOverrides?.wrath !== undefined
    ? Math.max(0, Number(manualOverrides.wrath ?? 0))
    : Math.max(0, wrath);

  const finalDamage = manualOverrides?.damage !== undefined
    ? manualOverrides.damage
    : damage;

  if (manualOverrides) {
    logDebug("WeaponDialog.computeFields: re-applying manual overrides", manualOverrides);
  }

  fields.pool = finalPool;
  fields.difficulty = finalDifficulty;
  fields.damage = finalDamage;
  fields.ed = finalEd;
  fields.ap = finalAp;
  fields.wrath = finalWrath;

  console.log("CE: Final values - pool:", fields.pool, "difficulty:", fields.difficulty, "damage:", fields.damage, "aim:", fields.aim);
  console.log("=== applyCombatExtender END ===");

  const actorForSafety = dialog.actor ?? dialog.token?.actor ?? null;
  const isEngagedForSafety = Boolean(getEngagedEffect(actorForSafety));
  const engagedRangedForSafety = Boolean(weapon?.isRanged && isEngagedForSafety);

  const hasAnyCombatOption = combatOptionsActive(fields);

  logDebug("WeaponDialog.computeFields: baseline vs final after CE", {
    baselinePool: systemBaselineSnapshot.pool,
    finalPool: fields.pool,
    hasAnyCombatOption,
    engagedRangedForSafety,
    manualOverrides
  });

  if (!dialog._combatOptionsManualOverrides &&
      !engagedRangedForSafety &&
      !hasAnyCombatOption &&
      typeof systemBaselineSnapshot.pool === "number") {
    fields.pool = Number(systemBaselineSnapshot.pool);
    fields.difficulty = Number(systemBaselineSnapshot.difficulty ?? fields.difficulty);
    fields.damage = systemBaselineSnapshot.damage;
    fields.ed = foundry.utils.deepClone(systemBaselineSnapshot.ed ?? fields.ed);
    fields.ap = foundry.utils.deepClone(systemBaselineSnapshot.ap ?? fields.ap);
  }

  return dialog.fields;
}

Hooks.once("ready", () => {
  if (game.wng?.registerScript) {
    game.wng.registerScript("dialog", {
      id: "wng-combat-extender",
      label: "Combat Extender",
      hide: () => false,
      submit(dialog) {
        dialog.flags = dialog.flags ?? {};
        dialog.flags.combatExtender = {
          delta: dialog._combatExtenderDelta ?? null
        };
      }
    });
  }
});

function trackManualOverrideSnapshots(app, html) {
  const $html = html instanceof jQuery ? html : $(html);

  const manualFieldSelectors = [
    'input[name="pool"]',
    'input[name="difficulty"]',
    'input[name="damage"]',
    'input[name="ed.value"]',
    'input[name="ed.dice"]',
    'input[name="ap.value"]',
    'input[name="ap.dice"]',
    'input[name="wrath"]'
  ];

  $html.off(".combatOptionsManual");
  $html.on(`change.combatOptionsManual input.combatOptionsManual`, manualFieldSelectors.join(","), (ev) => {
    const el = ev.currentTarget;
    const name = el.name;
    const fields = app.fields ?? (app.fields = {});

    const value = el.type === "number" ? Number(el.value ?? 0) : el.value;
    foundry.utils.setProperty(fields, name, value);

    const manualSnapshot = foundry.utils.deepClone(app._combatOptionsManualOverrides ?? {});

    if (name === "pool") {
      manualSnapshot.pool = fields.pool;
    } else if (name === "difficulty") {
      manualSnapshot.difficulty = fields.difficulty;
    } else if (name === "damage") {
      manualSnapshot.damage = fields.damage;
    } else if (name.startsWith("ed.")) {
      const manualEd = foundry.utils.deepClone(manualSnapshot.ed ?? {});
      manualEd.value = Number(fields.ed?.value ?? 0);
      manualEd.dice = Number(fields.ed?.dice ?? 0);
      manualSnapshot.ed = manualEd;
    } else if (name.startsWith("ap.")) {
      const manualAp = foundry.utils.deepClone(manualSnapshot.ap ?? {});
      manualAp.value = Number(fields.ap?.value ?? 0);
      manualAp.dice = Number(fields.ap?.dice ?? 0);
      manualSnapshot.ap = manualAp;
    } else if (name === "wrath") {
      manualSnapshot.wrath = Number(fields.wrath ?? 0);
    }

    const hasManualOverrides = Object.keys(manualSnapshot).length > 0;
    app._combatOptionsManualOverrides = hasManualOverrides
      ? foundry.utils.deepClone(manualSnapshot)
      : null;

    logDebug("WeaponDialog: manual override snapshot updated", {
      field: name,
      manualOverrides: app._combatOptionsManualOverrides
    });
  });
}

// ============================================================================
// PRIMARY HOOK: renderWeaponDialog
// ============================================================================
// FIX #1: Added render guard with app._isRendering flag
// FIX #2: Fixed cover override logic to only reset when target changes
// FIX #3: Early return in change handler to avoid double-render
// ============================================================================
Hooks.on("renderWeaponDialog", async (app, html) => {
  try {
    if (game.system.id !== "wrath-and-glory") return;

    // FIX #1: Prevent re-entrant rendering with guard flag
    if (app._isRendering) {
      logDebug("CE: Skipping render - already in progress");
      return;
    }
    app._isRendering = true;

    try {
      const wasJustPatched = ensureWeaponDialogPatched(app);

      const $html = html instanceof jQuery ? html : $(html);

      // If we just patched, fix the initial field values directly
      // (computeFields already ran before we patched, so values are wrong)
      if (wasJustPatched) {
        const actor = app.actor ?? app.token?.actor;
        const isEngaged = Boolean(getEngagedEffect(actor));
        const weapon = app.weapon;
        const traits = weapon?.system?.traits;
        const hasPistol = Boolean(traits?.has?.("pistol") || traits?.get?.("pistol"));
        const rangeBand = String(app.fields?.range ?? "").toLowerCase();

        if (isEngaged && weapon?.isRanged && hasPistol) {
          console.log("CE: Fixing initial values for first open (engaged)");
          
          // Clear aim in fields
          if (app.fields.aim) app.fields.aim = false;
          
          // Add +2 DN for engagement
          const newDifficulty = (app.fields.difficulty ?? 0) + 2;
          app.fields.difficulty = newDifficulty;
          $html.find('input[name="difficulty"]').val(newDifficulty);
          
          // Suppress short range bonus
          if (rangeBand === "short") {
            const newPool = (app.fields.pool ?? 0) - 1;
            app.fields.pool = newPool;
            $html.find('input[name="pool"]').val(newPool);
          }
        }
      }

      $html.find('.form-group').has('input[name="aim"]').remove();
      $html.find('.form-group').has('input[name="charging"]').remove();
      $html.find('.form-group').has('select[name="calledShot.size"]').remove();


      // Cache system-determined range band for computeFields / applyCombatExtender
      app._combatExtenderRangeBand = String($html.find('select[name="range"]').val() ?? "").toLowerCase();
      
      const attackSection = $html.find(".attack");
      if (!attackSection.length) return;

      const salvoValue = Number(app.weapon?.system?.salvo ?? app.weapon?.salvo ?? 0);
      const canPinning = Boolean(app.weapon?.isRanged) && Number.isFinite(salvoValue) && salvoValue > 1;

      const targetResolve = getTargetResolve(app);
      const normalizedResolve = Number.isFinite(targetResolve) ? Math.max(0, Math.round(targetResolve)) : null;
      const ctx = {
        open: app._combatOptionsOpen ?? false,
        isMelee: !!app.weapon?.isMelee,
        isRanged: !!app.weapon?.isRanged,
        hasHeavy: !!app.weapon?.system?.traits?.has?.("heavy"),
        canPinning,
        pinningResolve: normalizedResolve,
        fields: foundry.utils.duplicate(app.fields ?? {}),
        labels: {
          allOutAttack: COMBAT_OPTION_LABELS.allOutAttack,
          charge: COMBAT_OPTION_LABELS.charge,
          brace: COMBAT_OPTION_LABELS.brace,
          pinning: COMBAT_OPTION_LABELS.pinning,
          cover: "Cover",
          vision: "Vision",
          size: "Target Size",
          calledShot: "Called Shot",
          calledShotSize: "Target Size",
          disarm: COMBAT_OPTION_LABELS.calledShotDisarm,
          disarmNoteHeading: COMBAT_OPTION_LABELS.disarmNoteHeading,
          disarmNote: COMBAT_OPTION_LABELS.disarmNote
        },
        coverOptions: [
          { name: "cover", value: "", label: "No Cover" },
          { name: "cover", value: "half", label: "Half Cover (+1 DN)" },
          { name: "cover", value: "full", label: "Full Cover (+2 DN)" }
        ],
        visionOptions: [
          { name: "visionPenalty", value: "", label: "Normal" },
          { name: "visionPenalty", value: "twilight", label: "Twilight (+1 DN Ranged)" },
          { name: "visionPenalty", value: "dim", label: "Dim Light (+2 DN Ranged / +1 DN Melee)" },
          { name: "visionPenalty", value: "heavy", label: "Heavy Fog (+3 DN Ranged / +2 DN Melee)" },
          { name: "visionPenalty", value: "darkness", label: "Darkness (+4 DN Ranged / +3 DN Melee)" }
        ],
        sizeOptions: [
          { name: "sizeModifier", value: "", label: "Average Target (No modifier)" },
          { name: "sizeModifier", value: "tiny", label: "Tiny Target (+2 DN)" },
          { name: "sizeModifier", value: "small", label: "Small Target (+1 DN)" },
          { name: "sizeModifier", value: "large", label: "Large Target (+1 Die)" },
          { name: "sizeModifier", value: "huge", label: "Huge Target (+2 Dice)" },
          { name: "sizeModifier", value: "gargantuan", label: "Gargantuan Target (+3 Dice)" }
        ],
        calledShotSizes: [
          { value: "", label: "" },
          { value: "tiny", label: game.i18n.localize("SIZE.TINY") },
          { value: "small", label: game.i18n.localize("SIZE.SMALL") },
          { value: "medium", label: game.i18n.localize("SIZE.MEDIUM") }
        ]
      };

      const actor = app.actor ?? app.token?.actor;
      const fields = app.fields ?? (app.fields = {});
      let shouldRecompute = false;

      let canPistolsInMelee = app._combatOptionsCanPistolsInMelee;
      if (typeof canPistolsInMelee !== "boolean") {
        const pistolTrait = app.weapon?.system?.traits;
        const hasPistolTrait = Boolean(pistolTrait?.has?.("pistol") || pistolTrait?.get?.("pistol"));
        const isEngaged = Boolean(getEngagedEffect(actor));
        canPistolsInMelee = hasPistolTrait && isEngaged;
      }
      canPistolsInMelee = Boolean(canPistolsInMelee);

      const pistolsInMeleeInput = $html.find('input[name="pistolsInMelee"]');
      if (pistolsInMeleeInput.length) {
        pistolsInMeleeInput.prop("disabled", !canPistolsInMelee);
        if (!canPistolsInMelee) {
          if (foundry.utils.getProperty(fields, "pistolsInMelee")) {
            shouldRecompute = true;
          }
          pistolsInMeleeInput.prop("checked", false);
          foundry.utils.setProperty(fields, "pistolsInMelee", false);
        }
      }

      const disableAllOutAttack = Boolean(actor?.statuses?.has?.("full-defence"));
      const previousAllOutAttack = foundry.utils.getProperty(fields, "allOutAttack");

      if (disableAllOutAttack) {
        foundry.utils.setProperty(ctx.fields, "allOutAttack", false);
        foundry.utils.setProperty(fields, "allOutAttack", false);
        if (previousAllOutAttack) {
          shouldRecompute = true;
        }
      }

      ctx.disableAllOutAttack = disableAllOutAttack;

      // FIX #2: Only reset cover override when target changes, not when value equals default
      const currentTargetId = getTargetIdentifier(app);
      if (app._combatOptionsCoverTargetId !== currentTargetId) {
        app._combatOptionsCoverOverride = false;
        app._combatOptionsCoverTargetId = currentTargetId;
      }

      if (app._combatOptionsPinningResolve !== normalizedResolve) {
        app._combatOptionsPinningResolve = normalizedResolve;
        shouldRecompute = true;
      }

      const defaultCover = "";
      const normalizedDefaultCover = defaultCover ?? "";
      app._combatOptionsDefaultCover = defaultCover;

      if (!app._combatOptionsCoverOverride) {
        const previousCover = (foundry.utils.getProperty(fields, "cover") ?? "");
        if (previousCover !== normalizedDefaultCover) {
          shouldRecompute = true;
        }
        ctx.fields.cover = normalizedDefaultCover;
        foundry.utils.setProperty(fields, "cover", normalizedDefaultCover);
      }

      const defaultSize = app._combatOptionsDefaultSizeModifier ?? getTargetSize(app);
      app._combatOptionsDefaultSizeModifier = defaultSize;
      const defaultFieldValue = defaultSize === "average" ? "" : defaultSize;
      const previousSizeModifier = ctx.fields.sizeModifier ?? "";
      if (app._combatOptionsSizeOverride && previousSizeModifier === defaultFieldValue) {
        app._combatOptionsSizeOverride = false;
      }
      if (!app._combatOptionsSizeOverride) {
        if (previousSizeModifier !== defaultFieldValue) {
          shouldRecompute = true;
        }
        ctx.fields.sizeModifier = defaultFieldValue;
        foundry.utils.setProperty(fields, "sizeModifier", defaultFieldValue);
      }

      if (!canPinning) {
        foundry.utils.setProperty(ctx.fields, "pinning", false);
      }

      const existing = attackSection.find("[data-co-root]");
      const htmlFrag = await renderTemplate(`${TEMPLATE_BASE_PATH}/combat-options.hbs`, ctx);
      if (existing.length) {
        existing.replaceWith(htmlFrag);
      } else {
        const hr = attackSection.find('hr').first();
        if (hr.length) {
          hr.before(htmlFrag);
        } else {
          attackSection.append(htmlFrag);
        }
      }

      const root = attackSection.find("[data-co-root]");
      if (root.length && typeof app._onFieldChange === "function") {
        root.find("[name]").each((_, el) => {
          if (el.dataset?.co) return;

          const $el = $(el);
          $el.off(".wngCE");
          $el.on("change.wngCE", (ev) => app._onFieldChange(ev));
        });
      }

      if (!canPinning) {
        foundry.utils.setProperty(fields, "pinning", false);
      }

      root.off(".combatOptions");
      $html.off("change.combatOptions");

      root.on("toggle.combatOptions", () => {
        app._combatOptionsOpen = root.prop("open");
      });

      root.on("change.combatOptions", "input[data-co], select[data-co]", async (ev) => {
        const el = ev.currentTarget;
        const name = el.name;
        const value = el.type === "checkbox" ? el.checked : el.value;

        logDebug("CE change:", name, value);

        if (!name) {
          logError("Combat option control missing name attribute", { tagName: el.tagName, type: el.type });
          return;
        }

        if (name === "allOutAttack" && disableAllOutAttack) {
          const forcedValue = false;
          root.find('input[name="allOutAttack"]').prop("checked", forcedValue);
          foundry.utils.setProperty(app.fields ?? (app.fields = {}), name, forcedValue);
          foundry.utils.setProperty(app.userEntry ?? (app.userEntry = {}), name, forcedValue);
          return;
        }

        if (name === "sizeModifier") {
          app._combatOptionsSizeOverride = true;
        }

        if (name === "cover") {
          app._combatOptionsCoverOverride = true;
        }

        foundry.utils.setProperty(app.fields ?? (app.fields = {}), name, value);
        foundry.utils.setProperty(app.userEntry ?? (app.userEntry = {}), name, value);

        if (name === "calledShot.enabled") {
          root.find(".combat-options__called-shot").toggleClass("is-hidden", !value);
        }

        if (name === "allOutAttack" && !disableAllOutAttack) {
          await syncAllOutAttackCondition(actor, Boolean(value));
        }

        // FIX #3: _onFieldChange already calls render, but DON'T call it if we're
        // already in a render cycle (this would cause infinite loop)
        if (typeof app._onFieldChange === "function" && !app._isRendering) {
          await app._onFieldChange(ev);
          // Early return - don't do anything else after _onFieldChange
          return;
        }
      });

      trackManualOverrideSnapshots(app, $html);
      syncDialogInputsFromFields(app, $html);

    } finally {
      // Always clear the rendering flag
      app._isRendering = false;
    }

  } catch (err) {
    app._isRendering = false;
    logError("Failed to render combat options", err);
    console.error(err);
  }
});
