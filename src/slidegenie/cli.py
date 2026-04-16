"""CLI commands for slidegenie."""
import json
import click


@click.group()
def cli():
    """SlideGenie - Local AI slide generation tool."""
    pass


@cli.command()
@click.option("--prompt", "-p", required=True, help="Content description for the slide")
@click.option("--output", "-o", default="output.pptx", help="Output file path")
@click.option(
    "--mode", "-m",
    type=click.Choice(["image", "editable", "both"]),
    default="editable",
    help="Output mode: image (PNG only), editable (PPTX with editable shapes), both",
)
@click.option(
    "--make-type", "-t",
    type=click.Choice(["graphic", "flow", "matrix"]),
    default=None,
    help="Slide type (auto-select if omitted)",
)
def generate(prompt, output, mode, make_type):
    """Generate a single slide from a text prompt."""
    from slidegenie.pipeline import generate_slide

    result = generate_slide(
        prompt=prompt,
        output_path=output,
        mode=mode,
        make_type=make_type,
    )

    for key, path in result.items():
        click.echo(f"{key}: {path}")


@cli.command()
@click.option("--input", "-i", "input_path", required=True, help="Input image path")
@click.option("--output", "-o", default="output.pptx", help="Output PPTX path")
def convert(input_path, output):
    """Convert an existing image to editable PPTX."""
    from slidegenie.pipeline import convert_image

    result = convert_image(input_path=input_path, output_path=output)
    click.echo(f"pptx_path: {result['pptx_path']}")


@cli.command()
@click.option("--config", "-c", required=True, help="JSON config file with slide specs")
@click.option("--output", "-o", required=True, help="Output PPTX path")
@click.option(
    "--mode", "-m",
    type=click.Choice(["image", "editable", "both"]),
    default="editable",
    help="Output mode",
)
@click.option("--workers", "-w", default=4, help="Max parallel workers")
@click.option("--image-dir", default=None, help="Directory to save slide images")
def batch(config, output, mode, workers, image_dir):
    """Generate multiple slides in parallel into a single PPTX.

    Config JSON format:
    [
      {"prompt": "...", "make_type": "graphic"},
      {"prompt": "...", "make_type": "flow"}
    ]
    """
    from slidegenie.pipeline import batch_generate

    with open(config, encoding="utf-8") as f:
        specs = json.load(f)

    click.echo(f"Generating {len(specs)} slides in parallel (workers={workers})...")
    result = batch_generate(
        specs=specs,
        output_path=output,
        mode=mode,
        max_workers=workers,
        image_dir=image_dir,
    )
    click.echo(f"output: {result}")


@cli.command()
@click.option("--inputs", "-i", multiple=True, required=True, help="Input PPTX files to merge")
@click.option("--output", "-o", required=True, help="Output merged PPTX path")
def merge(inputs, output):
    """Merge multiple PPTX files into one (correctly copies embedded images)."""
    from slidegenie.pipeline import merge_pptx

    click.echo(f"Merging {len(inputs)} files...")
    result = merge_pptx(input_paths=list(inputs), output_path=output)
    click.echo(f"output: {result}")
