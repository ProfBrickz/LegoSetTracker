CREATE TYPE "public"."set_piece_type" AS ENUM('normal', 'minifig', 'extra', 'counterpart');

--> statement-breakpoint
CREATE TABLE "lego_colors" (
	"id" "smallserial" PRIMARY KEY NOT NULL,
	"bricklink_id" smallint NOT NULL,
	"bricklink_name" varchar(64) NOT NULL,
	"lego_id" smallint,
	"lego_name" varchar(64),
	CONSTRAINT "lego_colors_bricklink_id_unique" UNIQUE("bricklink_id")
);

--> statement-breakpoint
CREATE TABLE "lego_pieces" (
	"id" serial PRIMARY KEY NOT NULL,
	"bricklink_id" varchar(16) NOT NULL,
	"color_id" smallint,
	"bricklink_name" varchar(512) NOT NULL,
	"bricklink_category" varchar(256) NOT NULL
);

--> statement-breakpoint
CREATE TABLE "lego_set_pieces" (
	"id" serial PRIMARY KEY NOT NULL,
	"lego_set_id" integer NOT NULL,
	"piece_id" integer NOT NULL,
	"set_piece_type" "set_piece_type" NOT NULL,
	"amount_needed" smallint NOT NULL,
	"amount_Found" smallint NOT NULL
);

--> statement-breakpoint
CREATE TABLE "lego_set_themes" (
	"id" "smallserial" PRIMARY KEY NOT NULL,
	"bricklink_id" varchar(8) NOT NULL,
	"bricklink_name" varchar(64) NOT NULL,
	CONSTRAINT "lego_set_themes_bricklink_id_unique" UNIQUE("bricklink_id")
);

--> statement-breakpoint
CREATE TABLE "lego_sets" (
	"id" serial PRIMARY KEY NOT NULL,
	"set_number" varchar(16) NOT NULL,
	"name" varchar(256) NOT NULL,
	"theme_id" smallint NOT NULL,
	"release_year" smallint NOT NULL,
	"piece_count" smallint DEFAULT 0 NOT NULL,
	"minifig_count" smallint DEFAULT 0 NOT NULL,
	"lego_set_count" smallint DEFAULT 1 NOT NULL,
	CONSTRAINT "lego_sets_set_number_unique" UNIQUE("set_number")
);

--> statement-breakpoint
ALTER TABLE
	"lego_pieces"
ADD
	CONSTRAINT "lego_pieces_color_id_lego_colors_id_fk" FOREIGN KEY ("color_id") REFERENCES "public"."lego_colors"("id") ON DELETE no action ON UPDATE no action;

--> statement-breakpoint
ALTER TABLE
	"lego_set_pieces"
ADD
	CONSTRAINT "lego_set_pieces_lego_set_id_lego_sets_id_fk" FOREIGN KEY ("lego_set_id") REFERENCES "public"."lego_sets"("id") ON DELETE no action ON UPDATE no action;

--> statement-breakpoint
ALTER TABLE
	"lego_set_pieces"
ADD
	CONSTRAINT "lego_set_pieces_piece_id_lego_pieces_id_fk" FOREIGN KEY ("piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;

--> statement-breakpoint
ALTER TABLE
	"lego_sets"
ADD
	CONSTRAINT "lego_sets_theme_id_lego_set_themes_id_fk" FOREIGN KEY ("theme_id") REFERENCES "public"."lego_set_themes"("id") ON DELETE no action ON UPDATE no action;
