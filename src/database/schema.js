// Imports
import { integer, pgEnum, pgTable, serial, smallint, smallserial, varchar } from "drizzle-orm/pg-core";

// Enums
export const legoSetPieceDBType = pgEnum("set_piece_type", ["normal", "minifig", "extra", "counterpart"]);


// Database tables
export const legoColorsDBTable = pgTable("lego_colors", {
   databaseId: smallserial("id").primaryKey(),
   bricklinkId: smallint("bricklink_id").unique().notNull(),
   bricklinkName: varchar("bricklink_name", { length: 64 }).notNull(),
   legoId: smallint("lego_id"),
   legoName: varchar("lego_name", { length: 64 })
});

export const legoSetThemesDBTable = pgTable("lego_set_themes", {
   databaseId: smallserial("id").primaryKey(),
   bricklinkId: varchar("bricklink_id", { length: 8 }).unique().notNull(),
   bricklinkName: varchar("bricklink_name", { length: 64 }).notNull()
});

export const legoPiecesDBTable = pgTable("lego_pieces", {
   databaseId: serial("id").primaryKey(),
   bricklinkId: varchar("bricklink_id", { length: 16 }).notNull(),
   colorId: smallint("color_id").notNull().references(() => legoColorsDBTable.databaseId),
   name: varchar("bricklink_name", { length: 512 }).notNull(),
   category: varchar("category", { length: 256 }).notNull()
});

export const legoSetsDBTable = pgTable("lego_sets", {
   databaseId: serial("id").primaryKey(),
   setNumber: varchar("set_number", { length: 16 }).unique().notNull(),
   name: varchar("name", { length: 256 }).notNull(),
   themeId: smallint("theme_id").notNull().references(() => legoSetThemesDBTable.databaseId),
   yearReleased: smallint("year_released").notNull(),
   pieceCount: smallint("piece_count").notNull().default(0),
   minifigCount: smallint("minifig_count").notNull().default(0),
   setCount: smallint("set_count").notNull().default(1)
});

export const legoSetPiecesDBTable = pgTable("lego_set_pieces", {
   databaseId: serial("id").primaryKey(),
   legoSetId: integer("lego_set_id").notNull().references(() => legoSetsDBTable.databaseId),
   pieceId: integer("piece_id").notNull().references(() => legoPiecesDBTable.databaseId),
   setPieceType: legoSetPieceDBType("set_piece_type").notNull(),
   amountNeeded: smallint("amount_needed").notNull(),
   amountFound: smallint("amount_Found").notNull()
});
