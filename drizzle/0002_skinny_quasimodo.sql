CREATE TABLE "component_lego_pieces" (
	"component_piece_id" integer NOT NULL,
	"compound_piece_id" integer NOT NULL
);
--> statement-breakpoint
ALTER TABLE "component_lego_pieces" ADD CONSTRAINT "component_lego_pieces_component_piece_id_lego_pieces_id_fk" FOREIGN KEY ("component_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;--> statement-breakpoint
ALTER TABLE "component_lego_pieces" ADD CONSTRAINT "component_lego_pieces_compound_piece_id_lego_pieces_id_fk" FOREIGN KEY ("compound_piece_id") REFERENCES "public"."lego_pieces"("id") ON DELETE no action ON UPDATE no action;
