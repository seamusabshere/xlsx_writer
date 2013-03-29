# xlsx_writer

Writes (doesn't read or modify) XLSX files.

## Real-world usage

<p><a href="http://brighterplanet.com"><img src="https://s3.amazonaws.com/static.brighterplanet.com/assets/logos/flush-left/inline/green/rasterized/brighter_planet-160-transparent.png" alt="Brighter Planet logo"/></a></p>

We use `xlsx_writer` for [sustainability analytics at Brighter Planet](http://brighterplanet.com/case_studies).

## Credit

Based on the [original simple\_xlsx\_writer gem](https://github.com/harvesthq/simple_xlsx_writer) and [patches by mumboe](https://github.com/mumboe/simple_xlsx_writer)

Then I tore it down and rebuilt it:

* no longer constructs everything in a single zipstream... instead writes the individual files to /tmp and then zips them together
* absolute minimum XML - went through every line, testing to see if I could remove it
* no more block format - this was more appropriate when it was constructed as a zipstream

Features not present in simple_xlsx_writer:

* opinionated, non-customizable styles - Arial 10pt, left-aligned text and dates, right-aligned numbers and currency
* autofilter based on a cell range
* header and footer, with support for images (.emf only) and page numbers
* fits columns to text

## Wishlist

1. Optional shared string optimizer

## Example

    require 'xlsx_writer'
    
    doc = XlsxWriter.new

    # show TRUE for true but a blank cell instead of FALSE
    doc.quiet_booleans!
    
    sheet1 = doc.add_sheet("People")

    # freeze pane underneath the first (header) row
    sheet1.freeze_top_left = 'A2'
    
    # DATA
    
    sheet1.add_row([
      "DoB",
      "Name",
      "Occupation",
      "Salary",
      "Citations",
      "Average citations per paper"
    ])
    sheet1.add_row([
      Date.parse("July 31, 1912"), 
      "Milton Friedman",
      "Economist / Statistician",
      {:type => :Currency, :value => 10_000},
      500_000,
      0.31
    ])
    sheet1.add_autofilter 'A1:E1'

    # FORMATTING

    doc.page_setup.top = 1.5
    doc.header.right.contents = 'Corporate Reporting'
    doc.footer.left.contents = 'Confidential'
    doc.footer.right.contents = :page_x_of_y
    
    # if you really need images in header/footer: do it in Excel, save, unzip the xlsx... get the .emf files, "cropleft" (if necessary), etc. from there

    left_header_image = doc.add_image('image1.emf', 118, 107)
    left_header_image.croptop = '11025f'
    left_header_image.cropleft = '9997f'
    center_footer_image = doc.add_image('image2.emf', 116, 36)
    doc.header.left.contents = left_header_image
    doc.footer.center.contents = [ 'Powered by ', center_footer_image ]
    doc.page_setup.header = 0
    doc.page_setup.footer = 0

    # OUTPUT

    # You should move the file to where you want it
    require 'fileutils'
    ::FileUtils.mv doc.path, 'myfile.xlsx'

    # don't forget
    doc.cleanup

## Copyright

Copyright (c) 2012 Dee Zsombor, Justin Beck, Seamus Abshere. See LICENSE for details.
