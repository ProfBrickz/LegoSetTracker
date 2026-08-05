// Imports
import { integer, pgEnum, pgTable, serial, smallint, smallserial, varchar } from "drizzle-orm/pg-core";

// Enums
export const legoSetPieceDBType = pgEnum("set_piece_type", ["normal", "minifig", "minifigPieces", "extra", "counterpart"]);


// Database tables
export const legoColorsDBTable = pgTable("lego_colors", {
   databaseId: smallserial("id").primaryKey(),
   brickLinkId: smallint("bricklink_id").unique().notNull(),
   brickLinkName: varchar("bricklink_name", { length: 64 }).notNull(),
   legoId: smallint("lego_id"),
   legoName: varchar("lego_name", { length: 64 })
});

export const legoSetThemesDBTable = pgTable("lego_set_themes", {
   databaseId: smallserial("id").primaryKey(),
   brickLinkId: varchar("bricklink_id", { length: 8 }).unique().notNull(),
   brickLinkName: varchar("bricklink_name", { length: 64 }).notNull()
});

export const legoPiecesDBTable = pgTable("lego_pieces", {
   databaseId: serial("id").primaryKey(),
   brickLinkId: varchar("bricklink_id", { length: 16 }).notNull(),
   colorId: smallint("color_id").references(() => legoColorsDBTable.databaseId),
   brickLinkName: varchar("bricklink_name", { length: 512 }).notNull(),
   brickLinkCategory: varchar("bricklink_category", { length: 256 }).notNull()
});

export const stickeredLegoPiecesDBTable = pgTable("stickered_lego_pieces", {
   baseLegoPieceId: integer("base_piece_id").notNull().references(() => legoPiecesDBTable.databaseId),
   stickeredLegoPieceId: integer("stickered_piece_id").notNull().references(() => legoPiecesDBTable.databaseId)
});

export const componentLegoPiecesDBTable = pgTable("component_lego_pieces", {
   componentLegoPieceId: integer("component_piece_id").notNull().references(() => legoPiecesDBTable.databaseId),
   compoundLegoPieceId: integer("compound_piece_id").notNull().references(() => legoPiecesDBTable.databaseId)
});

export const legoSetsDBTable = pgTable("lego_sets", {
   databaseId: serial("id").primaryKey(),
   setNumber: varchar("set_number", { length: 16 }).unique().notNull(),
   name: varchar("name", { length: 256 }).notNull(),
   themeId: smallint("theme_id").notNull().references(() => legoSetThemesDBTable.databaseId),
   releaseYear: smallint("release_year").notNull(),
   pieceCount: smallint("piece_count").notNull().default(0),
   minifigCount: smallint("minifig_count").notNull().default(0),
   legoSetCount: smallint("lego_set_count").notNull().default(1)
});

export const legoSetPiecesDBTable = pgTable("lego_set_pieces", {
   databaseId: serial("id").primaryKey(),
   legoSetId: integer("lego_set_id").notNull().references(() => legoSetsDBTable.databaseId),
   legoPieceId: integer("piece_id").notNull().references(() => legoPiecesDBTable.databaseId),
   legoSetPieceType: legoSetPieceDBType("set_piece_type").notNull(),
   amountNeeded: smallint("amount_needed").notNull(),
   amountFound: smallint("amount_Found").notNull()
});
