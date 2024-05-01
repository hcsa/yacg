# Game rules

A game consists of multiple rounds. **The first player to win 2 rounds wins the game**.

## Playing a round

Each round starts with the players drawing 8 cards.

One of the players is A, the other one is B. In the first round, players decide who is A and who is B by flipping a
coin; for the following games, they alternate who’s A and who’s B.

Each round has 4 phases. On all of them, the players have options available to them. These options are chosen
alternately, always starting with player
A.

### Phase 1 - Opening phase

In this phase, effects cards are played. There are 3 types of effect cards:

- **Actions** (marked with <img src="./card_design/icons/card_action.svg" height=12>) in the top-right corner)  have an immediate effect, then go
  to the
  discard pile.
- **Fields** (marked with <img src="./card_design/icons/card_field.svg" height=12> in the top-right corner) stay in the field, and their effect
  applies
  throughout the round.
- **Auras** (marked with <img src="./card_design/icons/card_aura.svg" height=12> in the top-right corner) are attached
  to
  creatures.
    - It’s very unlikely you’ll get to play auras, since there aren’t creatures in play, but maybe it’ll happen with
      some mischievous effects…

**Cards are played to an effect stack**. When a stack is closed, cards are resolved from top (last played effect) to
bottom (first played effect), activating their effects sequentially.

There can be multiple effect stacks in this phase.

**Each player has the following options available to them**:

- **Hold**: do nothing. Can only be done when there are no cards in the stack.
- **Play an effect**: adds it to the stack.
- **Close the stack**: can only be done when there are cards in the stack. The stack's cards are then resolved.

The phase ends when both players choose consecutively to hold.

### Phase 2 - Creature phase

In this phase, creature cards are played. They are marked with <img src="./card_design/icons/card_creature.svg" height=12> in the top-right
corner.

Each player has the following options available to them:

- **Hold**: do nothing.
- **Play a creature**: place it on their side of the field.

The phase ends when both players choose consecutively to hold.

### Phase 3 - Development phase

This phase proceeds exactly the same as Phase 1.

### Phase 4 - Combat phase

This phase proceeds as follows:

- Player A chooses a <img src="./card_design/icons/card_creature.svg" height=12> from their side of the field.
- Player B chooses a <img src="./card_design/icons/card_creature.svg" height=12> from their side of the field.
- The chose creatures fight!
    - Each <img src="./card_design/icons/card_creature.svg" height=12> has <img src="./card_design/icons/stat_hp.svg" height=12> (Health Points, or
      HP), <img src="./card_design/icons/stat_atk.svg" height=12> (Attack, or Atk) and <img src="./card_design/icons/stat_spe.svg" height=12> (Speed, or Spe) as written on the
      card.
    - Both <img src="./card_design/icons/card_creature.svg" height=12> fight by attacking each other, reducing <img src="./card_design/icons/stat_hp.svg" height=12> of the
      opponent by an amount equal to their <img src="./card_design/icons/stat_atk.svg" height=12>. The order in which they attack depends on
      their <img src="./card_design/icons/stat_spe.svg" height=12>.
        - The <img src="./card_design/icons/card_creature.svg" height=12> with the highest <img src="./card_design/icons/stat_spe.svg" height=12> attacks first.
        - If they have the same <img src="./card_design/icons/stat_spe.svg" height=12>, then they attack simultaneously.
    - If <img src="./card_design/icons/stat_hp.svg" height=12> reaches 0, it's killed and put into the discard pile.
        - If a <img src="./card_design/icons/card_creature.svg" height=12> were to attack second but dies first, then its attack doesn't go
          through.
- This repeats over and over again, alternating the player who first chooses a <img src="./card_design/icons/card_creature.svg" height=12> to
  fight.
- **When a player has no more <img src="./card_design/icons/card_creature.svg" height=12> left, their opponent wins the round**.
    - If both players have no more creatures left, the round ends in a tie.

### Round aftermath

At the end of the round, **all cards on the field are put on they respective owner’s discard pile, and all cards on the
players’ hands return to their respective decks**. The decks are shuffled, and the remainder of the game is played with
the cards remaining on the deck.

---

## Energy

**To play a card, you need to pay energy**.

- The total amount of energy to be paid is in the top-left corner of the card.
- Some of the cost must be paid with energy of the same color as the card’s color. This is the number in the bottom-left
  corner of the card.
- The number in the bottom-right corner is the difference between the two. This cost can be paid with energy of any
  color.
- Brown cards are colorless. There is no brown energy, and their cost can be paid with energy of any color.

You earn energy as follows:

- At the start of the game, you choose 2 of the available plans to get energy. These can be:
    - Get energy at the beginning of every round
    - Get energy whenever you cast a <img src="./card_design/icons/card_creature.svg" height=12>
    - Get energy whenever an opponent <img src="./card_design/icons/card_creature.svg" height=12> is killed or destroyed.
    - Get energy whenever an ally <img src="./card_design/icons/card_creature.svg" height=12> is killed or destroyed.
    - Get energy whenever an ally <img src="./card_design/icons/card_creature.svg" height=12> survives until the end of the round.
    - Get energy whenever you lose a round.
- To each chosen plan, you attach a color. The 2 plans may have the same or different colors.
- Throughout the game, when a chosen plan’s condition applies, you get energy of the color attached to the plan.
- In addition, whenever a card from the deck enters your hand whose color matches one attached to any of the methods,
  you get +1 energy of that color.
    - This happens if you’re drawing the 8 cards at the start of a game, if you draw a card, if you search for a card
      from the deck to put it in your hand, etc.

> **Note: Energy carries over between rounds!**

---

## Misc details

### Deck-building

A deck has 50 cards, and there’s at most 3 copies of each card.

### Mulligan

When drawing the 8 cards at the start of a game, the player may choose to put their hand on the deck, shuffle again and
draw 8 cards. They don’t get energy from drawing the first 8 cards.

Every player can mulligan once per game.

### Buffs and debuffs

Actions and traits can influence a <img src="./card_design/icons/card_creature.svg" height=12>
’s <img src="./card_design/icons/stat_hp.svg" height=12>, <img src="./card_design/icons/stat_atk.svg" height=12> or <img src="./card_design/icons/stat_spe.svg" height=12>. When computing
a <img src="./card_design/icons/card_creature.svg" height=12>’s stat, buffs are added before debuffs.

None of these stats can go below 0; if that
were to happen, the stat's value is 0.

- If <img src="./card_design/icons/stat_hp.svg" height=12> is 0, the <img src="./card_design/icons/card_creature.svg" height=12> dies and is put into the discard pile. This
  may happen outside of combat.
- If <img src="./card_design/icons/stat_atk.svg" height=12> is 0, the <img src="./card_design/icons/card_creature.svg" height=12>'s attacks don't do damage.

### Standard speed

As a rule of thumb, the “standard” value for <img src="./card_design/icons/stat_spe.svg" height=12> is 2.

### Weather

Some cards or <img src="./card_design/icons/card_creature.svg" height=12>'s traits may set up weather: sun, rain or hail.

Weather lasts for 2 rounds (the one it’s set, plus the one after), or until weather is set up again.

### Combat stall

The combat phase has a max of 20 fights. If there are <img src="./card_design/icons/card_creature.svg" height=12> left after that, each player
adds up the <img src="./card_design/icons/stat_hp.svg" height=12> left on the creatures on their side. The round’s winner is the one with the highest
value; if the values are the same, the round ends in a tie.

### Empty deck

If a player’s deck runs out, they keep playing as normal.

- If they were to draw 2 cards but have no more cards left, they don’t draw any.
- If they were to start a round, and they only have 5 cards left, that’s all they draw.

If neither player has cards left at the beginning of a match, the game’s winner is the one that has won the most rounds
so far. If both players have won the same number of rounds, the game ends in a tie.

### Kill VS Destroy

A <img src="./card_design/icons/card_creature.svg" height=12> is "killed" when its HP reaches 0, either by taking combat damage or due to
effects or traits that reduce its HP. A <img src="./card_design/icons/card_creature.svg" height=12> is "destroyed" when an effect or trait
causes it to be destroyed.

Almost all effects or traits destroy a <img src="./card_design/icons/card_creature.svg" height=12>, rather than killing it. If it's not clear in
which category it falls, assume it destroys.

### Targets

Some cards often refer to "targets". When a card is played, any targets it requires must be declared by whoever played
it, following any restrictions on the targets.

> Example: If a card mentions a "target ally <img src="./card_design/icons/card_creature.svg" height=12>", then the player must declare
> a <img src="./card_design/icons/card_creature.svg" height=12> on its side to take the effect of the card. If there are no
> ally <img src="./card_design/icons/card_creature.svg" height=12>, then the card can't be played.

The target restrictions apply when the card is played, and don't apply when the card is resolved.

A card's targets may no longer be on the field after an effect is played but before it's resolved. The card doesn't
change targets if this happens. When the card is resoled, changes to targets that aren't on the field do not apply,
while changes to targets that are on the field still apply.