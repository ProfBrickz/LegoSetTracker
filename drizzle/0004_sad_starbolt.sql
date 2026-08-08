CREATE TYPE "public"."lego_set_status_type" AS ENUM('scraping', 'done');--> statement-breakpoint
ALTER TYPE "public"."set_piece_type" RENAME TO "lego_set_piece_type";--> statement-breakpoint
ALTER TABLE "lego_sets" ADD COLUMN "status" "lego_set_status_type" NOT NULL;
