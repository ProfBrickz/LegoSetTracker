CREATE TABLE "stickered_lego_pieces" (
	"base_piece_id" integer NOT NULL,
	"stickered_piece_id" integer NOT NULL
);
--> statement-breakpoint
ALTER TABLE "stickered_lego_pieces" ADD CONSTRAINT "stickered_lego_pieces_base_piece_id_lego_pieces_id_fk" FOREIGN KEY ("base_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "stickered_lego_pieces" ADD CONSTRAINT "stickered_lego_pieces_stickered_piece_id_lego_pieces_id_fk" FOREIGN KEY ("stickered_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;
--> statement-breakpoint
-- Add stickered lego pieces
INSERT INTO stickered_lego_pieces (base_piece_id, stickered_piece_id)
SELECT
	base_lego_piece.id AS base_piece_id,
	stickered_lego_piece.id AS stickered_piece_id
FROM lego_pieces stickered_lego_piece
JOIN lego_pieces base_lego_piece ON
    base_lego_piece.bricklink_id = split_part(stickered_lego_piece.bricklink_id, 'pb', 1)
WHERE
	stickered_lego_piece.bricklink_name ILIKE '%(Sticker)%'
	AND stickered_lego_piece.color_id  = base_lego_piece.color_id;
