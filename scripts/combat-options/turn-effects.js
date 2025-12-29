import {
  MODULE_ID,
  PERSISTENT_DAMAGE_CONDITIONS,
  SLOWED_CONDITIONS
} from "./constants.js";
import { log, logError } from "./logging.js";
import { isActivePrimaryGM } from "./permissions.js";

export async function syncAllOutAttackCondition(actor, enabled) {
  if (!actor || game.system?.id !== "wrath-and-glory") return;
  if (typeof actor.hasCondition !== "function") return;

  if (enabled) {
    if (actor.hasCondition("all-out-attack")) return;
    if (typeof actor.addCondition !== "function") return;
    try {
      await actor.addCondition("all-out-attack", { [MODULE_ID]: { source: "combat-options" } });
    } catch (err) {
      logError("Failed to add All-Out Attack condition", err);
    }
    return;
  }

  if (typeof actor.removeCondition !== "function") return;
  try {
    await actor.removeCondition("all-out-attack");
  } catch (err) {
    logError("Failed to remove All-Out Attack condition", err);
  }
}

export async function removeAllOutAttackFromActor(actor) {
  if (!actor || game.system?.id !== "wrath-and-glory") return;
  if (!actorHasStatus(actor, "all-out-attack")) return;

  if (typeof actor.removeCondition === "function") {
    try {
      await actor.removeCondition("all-out-attack");
    } catch (err) {
      logError("Failed to remove All-Out Attack condition", err);
    }
    return;
  }

  const effect = actor.effects?.find?.((entry) => entry?.statuses?.has?.("all-out-attack"));
  if (effect && typeof effect.delete === "function") {
    try {
      await effect.delete();
    } catch (err) {
      logError("Failed to delete All-Out Attack effect", err);
    }
  }
}

function actorHasStatus(actor, statusId) {
  if (!actor) return false;
  if (typeof actor.hasCondition === "function") {
    try {
      return actor.hasCondition(statusId);
    } catch (err) {
      logError("Failed to check condition", err);
    }
  }

  const statuses = actor.statuses ?? actor.effects?.reduce?.((acc, effect) => {
    if (!effect) return acc;
    if (effect.statuses) {
      for (const status of effect.statuses) {
        acc.add(status);
      }
    }
    return acc;
  }, new Set());

  return statuses?.has?.(statusId) ?? false;
}

function sanitizePersistentDamageValue(value) {
  const num = Number(value);
  if (Number.isNaN(num) || !Number.isFinite(num)) return 0;
  return Math.max(0, Math.floor(num));
}

function extractPersistentDamageOverride(effect) {
  const systemId = game.system?.id;
  if (!effect || typeof effect.getFlag !== "function" || !systemId) return null;

  const flags = effect.getFlag(systemId, "value");
  if (flags !== undefined && flags !== null && flags !== "") {
    return flags;
  }

  if (effect.specifier !== undefined && effect.specifier !== null && effect.specifier !== "") {
    return effect.specifier;
  }

  return null;
}

async function evaluatePersistentDamage(effect, conditionConfig) {
  if (!effect || !conditionConfig) return null;

  const label = conditionConfig.labelKey ?
    game.i18n.localize(conditionConfig.labelKey) : (effect.name ?? conditionConfig.id);
  const override = extractPersistentDamageOverride(effect);
  let detail = null;
  let amount;

  const defaultConfig = conditionConfig.default ?? {};

  if (typeof override === "string") {
    const formula = override.trim();
    if (formula) {
      try {
        const roll = await (new Roll(formula)).evaluate({ async: true });
        amount = roll.total ?? 0;
        detail = { formula: roll.formula ?? formula, total: roll.total ?? 0 };
      } catch (err) {
        logError(`Failed to evaluate persistent damage formula "${formula}" for ${label}`, err);
        amount = defaultConfig.amount ?? 0;
      }
    }
  } else if (override !== null) {
    amount = override;
  }

  if (amount === undefined) {
    if (defaultConfig.formula) {
      try {
        const roll = await (new Roll(defaultConfig.formula)).evaluate({ async: true });
        amount = roll.total ?? 0;
        detail = { formula: roll.formula ?? defaultConfig.formula, total: roll.total ?? 0 };
      } catch (err) {
        logError(`Failed to evaluate persistent damage formula "${defaultConfig.formula}" for ${label}`, err);
        amount = defaultConfig.amount ?? 0;
      }
    } else {
      amount = defaultConfig.amount ?? 0;
    }
  }

  const normalizedAmount = sanitizePersistentDamageValue(amount);
  return {
    conditionId: conditionConfig.id,
    label,
    amount: normalizedAmount,
    detail,
    effect
  };
}

function shouldHandlePersistentDamage() {
  return game.system?.id === "wrath-and-glory" && isActivePrimaryGM();
}

const pendingPersistentDamageCombatants = new Map();

function isTurnChangeUpdate(changed) {
  if (!changed) return false;
  return ["turn", "combatantId", "round"].some((key) => Object.prototype.hasOwnProperty.call(changed, key));
}

function markPendingPersistentDamage(combat, changed) {
  if (!combat || !isTurnChangeUpdate(changed)) return;
  const currentCombatant = combat.combatant;
  if (!currentCombatant?.id) return;
  const actor = currentCombatant.actor;
  if (!actorHasPersistentDamage(actor)) return;
  pendingPersistentDamageCombatants.set(combat.id, currentCombatant.id);
}

function actorHasPersistentDamage(actor) {
  if (!actor || typeof actor.hasCondition !== "function") return false;
  return Object.values(PERSISTENT_DAMAGE_CONDITIONS).some((config) => actor.hasCondition(config.id));
}

function promptPendingPersistentDamage(combat) {
  if (!combat) return;
  const pendingId = pendingPersistentDamageCombatants.get(combat.id);
  if (!pendingId) return;
  pendingPersistentDamageCombatants.delete(combat.id);

  const previousCombatant = typeof combat.combatants?.get === "function"
    ? combat.combatants.get(pendingId)
    : combat.combatants?.find?.((c) => c?.id === pendingId);

  if (!previousCombatant) return;

  const maybePromise = promptPersistentDamageAtTurnEnd(previousCombatant);
  if (maybePromise?.catch) {
    maybePromise.catch((err) => logError("Failed to prompt for persistent damage", err));
  }
}

function cleanupPendingPersistentDamageForCombat(combatId) {
  if (!combatId) return;
  pendingPersistentDamageCombatants.delete(combatId);
}

function cleanupPendingPersistentDamageForCombatant(combatant) {
  const parentCombat = combatant?.parent;
  const combatId = parentCombat?.id;
  if (!combatId) return;

  const pendingId = pendingPersistentDamageCombatants.get(combatId);
  if (pendingId && pendingId === combatant.id) {
    pendingPersistentDamageCombatants.delete(combatId);
  }
}

async function promptPersistentDamageAtTurnEnd(combatant) {
  const actor = combatant?.actor;
  if (!actor || typeof actor.hasCondition !== "function") return;

  const entries = await Promise.all(
    Object.values(PERSISTENT_DAMAGE_CONDITIONS)
      .filter((config) => actor.hasCondition(config.id))
      .map(async (config) => {
        const effect = actor.effects?.find?.((entry) => entry?.statuses?.has?.(config.id)) ?? null;
        return evaluatePersistentDamage(effect, config);
      })
  );

  const results = entries.filter(Boolean);
  if (!results.length) return;

  const detailItems = results.map((entry) => {
    const detail = entry.detail;
    const tail = detail?.formula
      ? ` (${detail.formula}${detail.total !== undefined ? ` = ${detail.total}` : ""})`
      : "";
    return `<li><strong>${entry.label}</strong>: ${entry.amount}${tail}</li>`;
  }).join("");

  const summary = results.map((entry) => entry.amount).reduce((total, value) => total + value, 0);

  // Show dialog with adjustable damage input
  const result = await Dialog.wait({
    title: game.i18n.localize("WNGCE.PersistentDamage.Title"),
    content: `
      <p>${game.i18n.localize("WNGCE.PersistentDamage.Description")}</p>
      <ol>${detailItems}</ol>
      <p><strong>${game.i18n.format("WNGCE.PersistentDamage.Total", { total: summary })}</strong></p>
      <div class="form-group" style="margin-top: 1em;">
        <label style="font-weight: bold;">Mortal Wounds to Apply:</label>
        <input type="number" name="persistent-damage" value="${summary}" min="0" 
               style="width: 100%; margin-top: 0.25em;" />
        <p class="notes" style="margin-top: 0.5em; font-size: 0.9em; font-style: italic;">
          You can adjust this value to modify persistent damage.
        </p>
      </div>
    `,
    buttons: {
      apply: {
        icon: '<i class="fas fa-check"></i>',
        label: "Apply Damage",
        callback: (html) => {
          const $html = html instanceof jQuery ? html : $(html);
          const input = $html.find('[name="persistent-damage"]')[0];
          const value = input?.value ?? summary;
          return sanitizePersistentDamageValue(value);
        }
      },
      cancel: {
        icon: '<i class="fas fa-times"></i>',
        label: "Cancel",
        callback: () => null
      }
    },
    default: "apply",
    close: () => null,
    rejectClose: false
  });

  // Apply the damage if user didn't cancel and entered a valid amount
  if (result !== null && result > 0) {
    try {
      // Apply mortal wounds to the actor
      const report = await actor.applyDamage(0, { mortal: result });
      
      // Create chat message showing the damage was applied
      const activeTokens = typeof actor.getActiveTokens === "function" ? actor.getActiveTokens() : [];
      const token = activeTokens[0]?.document ?? actor.prototypeToken;
      
      await ChatMessage.create({
        content: `<p data-tooltip-direction="LEFT" data-tooltip="${report.breakdown}">${report.message}</p>`,
        speaker: ChatMessage.getSpeaker({ actor: actor, token: token })
      });
      
      log(`Applied ${result} mortal wounds from persistent damage to ${actor.name}`);
    } catch (err) {
      logError(`Failed to apply persistent damage to ${actor.name}`, err);
      ui.notifications.error(`Failed to apply persistent damage to ${actor.name}: ${err.message}`);
    }
  } else if (result === 0) {
    log(`Persistent damage for ${actor.name} was set to 0 and not applied`);
  } else {
    log(`Persistent damage application cancelled for ${actor.name}`);
  }
}

function shouldNotifySlowedConditions() {
  return game.system?.id === "wrath-and-glory" && isActivePrimaryGM();
}

function collectSlowedConditions(actor) {
  if (!actor) return [];

  const toCheck = Array.isArray(actor.conditions)
    ? actor.conditions
    : (actor.conditions?.contents ?? actor.conditions?.toObject?.() ?? []);

  if (!toCheck?.length) return [];

  const results = [];
  for (const entry of toCheck) {
    if (!entry) continue;
    const conditionId = entry.statuses ? entry.id : entry;
    if (!conditionId) continue;

    const match = SLOWED_CONDITIONS.find((config) => config.id === conditionId);
    if (!match) continue;

    const label = match.labelKey ? game.i18n.localize(match.labelKey) : (entry.name ?? match.id ?? "");
    results.push({ id: match.id, label });
  }

  return results;
}

function formatLocalizedList(items) {
  if (!items.length) return "";
  if (items.length === 1) return items[0];
  if (items.length === 2) {
    const conjunction = game?.i18n?.localize?.("WNG.Common.And") ?? "and";
    return `${items[0]} ${conjunction} ${items[1]}`;
  }

  const conjunction = game?.i18n?.localize?.("WNG.Common.And") ?? "and";
  const head = items.slice(0, -1).join(", ");
  const tail = items[items.length - 1];
  return `${head}, ${conjunction} ${tail}`;
}

function getSlowedNotificationRecipients(actor) {
  if (!actor) return [];

  const users = Array.isArray(game?.users) ? game.users : [];
  const canTestPermission = typeof actor.testUserPermission === "function";
  const recipients = new Set();

  for (const user of users) {
    if (!user) continue;
    if (user.isGM) {
      recipients.add(user.id);
      continue;
    }

    if (!canTestPermission) continue;
    try {
      if (actor.testUserPermission(user, "OWNER")) {
        recipients.add(user.id);
      }
    } catch (err) {
      logError("Failed to evaluate slowed condition permissions", err);
    }
  }

  return Array.from(recipients);
}

function notifySlowedConditions(combat) {
  if (!shouldNotifySlowedConditions()) return;
  const combatant = combat?.combatant;
  const actor = combatant?.actor;
  if (!actor) return;

  const slowedConditions = collectSlowedConditions(actor);
  if (!slowedConditions.length) return;

  const rawName = actor.name ?? combatant.name ?? game.i18n.localize("WNG.SlowedConditions.UnknownName");
  const safeName = foundry.utils.escapeHTML(rawName);
  const conditionLabels = slowedConditions.map((entry) => foundry.utils.escapeHTML(entry.label ?? entry.id));
  const conditionsText = formatLocalizedList(conditionLabels);

  const activeTokens = typeof actor.getActiveTokens === "function" ? actor.getActiveTokens() : [];
  const combatantToken = combatant.token ?? null;
  const tokenDocument = combatantToken ?? activeTokens[0]?.document ?? null;
  const tokenObject = combatantToken?.object ?? activeTokens[0] ?? null;
  const speakerData = {
    actor,
    token: tokenObject ?? tokenDocument ?? undefined,
    scene: combat?.scene?.id ?? tokenDocument?.parent?.id
  };

  const message = game.i18n.format("WNG.SlowedConditions.ChatReminder", {
    name: safeName,
    conditions: conditionsText
  });

  const recipients = getSlowedNotificationRecipients(actor);
  const messageData = {
    content: `<p>${message}</p>`,
    speaker: ChatMessage.getSpeaker(speakerData) ?? undefined,
    flags: {
      [MODULE_ID]: {
        slowedReminder: true,
        conditions: slowedConditions.map((entry) => entry.id)
      }
    }
  };

  if (recipients.length) {
    messageData.whisper = recipients;
  }

  const maybePromise = ChatMessage.create(messageData);
  if (maybePromise?.catch) {
    maybePromise.catch((err) => logError("Failed to create slowed reminder chat message", err));
  }
}

export function registerTurnEffectHooks() {
  Hooks.on("preUpdateCombat", (combat, changed) => {
    if (!shouldHandlePersistentDamage()) return;
    markPendingPersistentDamage(combat, changed);
  });

Hooks.on("combatTurn", async (combat) => {
  notifySlowedConditions(combat);

  // Persistent Damage: end-of-turn behavior (your existing pattern)
  if (!shouldHandlePersistentDamage()) return;
  promptPendingPersistentDamage(combat);
});


Hooks.on("updateCombat", (combat, changed) => {
  // Only run when the turn actually advanced
  if (!isTurnChangeUpdate(changed)) return;
  if (!isActivePrimaryGM()) return;

  // Defer so Combat has updated its active combatant getter
  setTimeout(async () => {
    const actor = combat?.combatant?.actor;
    if (actor) await removeAllOutAttackFromActor(actor);
  }, 0);

  // keep your existing PD fallback:
  if (shouldHandlePersistentDamage()) promptPendingPersistentDamage(combat);
});


  Hooks.on("deleteCombat", async (combat) => {
    cleanupPendingPersistentDamageForCombat(combat?.id);
    if (game.system?.id === "wrath-and-glory") {
      const combatants = combat?.combatants ?? [];
      for (const combatant of combatants) {
        await removeAllOutAttackFromActor(combatant?.actor);
      }
    }
  });

  Hooks.on("deleteCombatant", async (combatant) => {
    cleanupPendingPersistentDamageForCombatant(combatant);
    if (game.system?.id === "wrath-and-glory") {
      await removeAllOutAttackFromActor(combatant?.actor);
    }
  });

}

export { actorHasStatus };
