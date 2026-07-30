CREATE TABLE "stickered_lego_pieces" (
	"base_piece_id" integer NOT NULL,
	"stickered_piece_id" integer NOT NULL
);
--> statement-breakpoint
ALTER TABLE "stickered_lego_pieces" ADD CONSTRAINT "stickered_lego_pieces_base_piece_id_lego_pieces_id_fk" FOREIGN KEY ("base_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "stickered_lego_pieces" ADD CONSTRAINT "stickered_lego_pieces_stickered_piece_id_lego_pieces_id_fk" FOREIGN KEY ("stickered_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;
