#region Namespaces
using System;
using System.IO;
using Microsoft.Office.Core;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using PptType = Microsoft.Office.Interop.PowerPoint.PpPlaceholderType;
#endregion // Namespaces

namespace Ppt2txt
{
  class Program
  {
    /// <summary>
    /// Usage prompt.
    /// </summary>
    static string[] _usage = new string[] {
      "Ppt2txt 3.1 * Powerpoint Slide Deck Text Extractor",
      "  (C) 2007-2013 Jeremy Tammik Autodesk Inc.",
      "usage:  ppt2txt [-f textfilename] [-t] pptfilename",
			"  -t: skip adding prefix 'Title: ' to each slide title",
      "  -f: write output to textfilename.txt",
      "      default: pptfilename.txt",
      "      -f-: stdout"
    };

    /// <summary>
    /// Constant representing 'True' in Powerpoint API.
    /// </summary>
    const MsoTriState _t = MsoTriState.msoTrue;

    /// <summary>
    /// Windows carriage return + linefeed combination.
    /// </summary>
    const string _crlf = "\r\n";

    /// <summary>
    /// Prepend a linefeed to every
    /// non-empty input string.
    /// </summary>
    static string PrependLinefeed( string s )
    {
      return 0 == s.Length ? string.Empty : _crlf + s;
    }

    /// <summary>
    /// Append a new paragraph p to the given text s
    /// with a linefeed separator in between if s is
    /// not empty.
    /// </summary>
    static string AppendParagraph( string s, string p )
    {
      if( 0 < s.Length )
      {
        s += _crlf;
      }
      return s + p;
    }

    #region Escape unpleasant Powerpoint characters
    //static Encoding _ascii = Encoding.ASCII;

    /// <summary>
    /// Remove unpleasant non-ascii characters
    /// that crop up in powerpoint.
    /// </summary>
    static string normalise_string( string s )
    {
      s = s.Replace( "…", "..." );
      s = s.Replace( '–', '-' );
      s = s.Replace( '‘', '\'' ); // backward apostrophe from powerpoint
      s = s.Replace( '’', '\'' ); // forward apostrophe from powerpoint
      s = s.Replace( ' ', ' ' ); // strange space char from powerpoint
      //
      // quote:
      //
      s = s.Replace( '\u0060', '\'' ); // grave accent
      s = s.Replace( '\u00b4', '\'' ); // acute accent
      s = s.Replace( '\u2018', '\'' ); // left single quotation mark
      s = s.Replace( '\u2019', '\'' ); // right single quotation mark
      s = s.Replace( '\u201C', '"' );  // left double quotation mark
      s = s.Replace( '\u201D', '"' );  // right double quotation mark
      //
      // space:
      //
      s = s.Replace( '\x0b', '\n' ); // vertical tab
      s = s.Replace( '\u0020', ' ' ); // space basic latin
      s = s.Replace( '\u00a0', ' ' ); // no-break space latin-1 supplement
      s = s.Replace( '\u1680', ' ' ); // ogham space mark ogham
      s = s.Replace( '\u180e', ' ' ); // mongolian vowel separator mongolian
      s = s.Replace( '\u2000', ' ' ); // en quad general punctuation
      s = s.Replace( '\u2001', ' ' ); // em quad
      s = s.Replace( '\u2002', ' ' ); // en space
      s = s.Replace( '\u2003', ' ' ); // em space
      s = s.Replace( '\u2004', ' ' ); // three-per-em space
      s = s.Replace( '\u2005', ' ' ); // four-per-em space
      s = s.Replace( '\u2006', ' ' ); // six-per-em space
      s = s.Replace( '\u2007', ' ' ); // figure space
      s = s.Replace( '\u2008', ' ' ); // punctuation space
      s = s.Replace( '\u2009', ' ' ); // thin space
      s = s.Replace( '\u200a', ' ' ); // hair space
      s = s.Replace( '\u202f', ' ' ); // narrow no-break space
      s = s.Replace( '\u205f', ' ' ); // medium mathematical space
      s = s.Replace( '\u3000', ' ' ); // ideographic space

      return s;

      //using System.Globalization;
      //using System.Text;
      //System.Globalization.UnicodeCategory.
      //s = s.Replace( "", "-->" ); // right arrow
      //s = s.Replace( '', '\'' ); // backquote Unicode characters
      //return Encoding.ASCII.GetBytes( s ).ToString();
      //return _ascii.GetString( _ascii.GetBytes( s ) );
    }
    #endregion // Escape unpleasant Powerpoint characters

    /// <summary>
    /// Get all the text from a given Ppt
    /// shape, delimited by _crlf.
    /// </summary>
    static string GetShapeText( Ppt.Shape shape )
    {
      string s = null;

      if( _t == shape.HasTextFrame
        && _t == shape.TextFrame.HasText )
      {
        s = shape.TextFrame.TextRange.Text.Trim();

        string[] a = s.Split(
          new char[] { '\r', '\n' } );

        s = string.Empty;

        foreach( string s2 in a )
        {
          s += PrependLinefeed( s2.Trim() );
        }
        s = s.Trim();

        if( 0 == s.Length )
        {
          s = null;
        }
      }
      return s;
    }

    static int Main( string[] args )
    {
      // Command line argument values.

      string filename_in = null;
      string filename_out = string.Empty;
      bool title_prefix = true;

      // Process command line arguments.

      int n = args.Length;
      int i = 0;

      while( i < n )
      {
        string a = args[i];

        if( '-' == a[0] )
        {
          if( 't' == a[1] )
          {
            title_prefix = !title_prefix;
          }
          else if( 'f' == a[1] )
          {
            filename_out = a.Substring( 2 );

            if( 0 == filename_out.Length )
            {
              ++i;
              if( i >= n )
              {
                Console.Error.WriteLine(
                  "-f option lacks output filename" );

                break;
              }
              filename_out = args[i];
            }
          }
          else
          {
            Console.Error.WriteLine(
              string.Format( "invalid option '{0}'",
                a ) );

            break;
          }
        }
        else if( null == filename_in )
        {
          filename_in = a;

          if( !File.Exists( filename_in ) )
          {
            if( File.Exists( filename_in + ".ppt" ) )
            {
              filename_in = filename_in + ".ppt";
            }
            else if( File.Exists(
              filename_in + ".pptx" ) )
            {
              filename_in = filename_in + ".pptx";
            }
          }
          if( !File.Exists( filename_in ) )
          {
            Console.Error.WriteLine(
              string.Format(
                "unable to open input file '{0}'",
                filename_in ) );

            break;
          }
        }
        else
        {
          Console.Error.WriteLine(
            string.Format( "invalid argument '{0}'",
              a ) );

          break;
        }
        ++i;
      }

      if( null == filename_in
        || !File.Exists( filename_in ) )
      {
        foreach( string s in _usage )
        {
          Console.Error.WriteLine( s );
        }
        return 1;
      }

      // Determine full input and output filenames.

      filename_in = Path.GetFullPath( filename_in );

      string ext = "txt";

      if( 0 == filename_out.Length )
      {
        filename_out = Path.ChangeExtension(
          filename_in, ext );
      }
      else if( !filename_out.Equals( "-" ) )
      {
        filename_out = Path.Combine(
          Path.GetDirectoryName( filename_in ),
          filename_out + "." + ext );
      }

      // Open output file and process ppt input.

      using( StreamWriter sw
        = filename_out.Equals( "-" )
          ? new StreamWriter( Console.OpenStandardOutput() )
          : new StreamWriter( filename_out ) )
      {
        if( null == sw )
        {
          Console.Error.WriteLine(
            string.Format(
              "unable to write to output file '{0}'",
                filename_out ) );
        }
        else
        {
          string s, title, subtitle, body, notes;

          Ppt.Application app = new Ppt.Application();

          app.Visible = _t;

          Ppt._Presentation p = app.Presentations.Open(
            filename_in, _t, _t, _t );

          foreach( Ppt._Slide slide in p.Slides )
          {
            title = subtitle = body = string.Empty;

            foreach( Ppt.Shape shape in slide.Shapes )
            {
              s = GetShapeText( shape );

              if( null != s )
              {
                if( MsoShapeType.msoPlaceholder
                  == shape.Type )
                {
                  switch( shape.PlaceholderFormat.Type )
                  {
                    case PptType.ppPlaceholderTitle:
                    case PptType.ppPlaceholderCenterTitle:
                      title = AppendParagraph(
                        title, s );
                      break;

                    case PptType.ppPlaceholderSubtitle:
                      subtitle = AppendParagraph(
                        subtitle, s );
                      break;

                    case PptType.ppPlaceholderBody:
                    default: // e.g., ppPlaceholderObject
                      body = AppendParagraph(
                        body, s );
                      break;
                  }
                }
                else
                {
                  body = AppendParagraph(
                    body, s );
                }
              }
            }

            // Retrieve notes text.

            notes = string.Empty;

            foreach( Ppt.Shape shape
              in slide.NotesPage.Shapes )
            {
              s = GetShapeText( shape );

              if( null != s )
              {
                if( slide.SlideIndex.ToString() != s )
                {
                  notes = AppendParagraph(
                    notes, s );
                }
              }
            }

            // Write output for current slide.

            if( 0 < ( title.Length + subtitle.Length
              + body.Length + notes.Length ) )
            {
              s = ( ( 0 == title.Length )
                  ? ( _crlf + "Slide " + slide.SlideIndex.ToString() )
                  : ( _crlf + ( title_prefix ? "Title: " : "" ) + title ) )
                + PrependLinefeed( subtitle )
                + PrependLinefeed( body )
                + PrependLinefeed( notes );

              sw.WriteLine(
                normalise_string(
                  s.Replace( "\n", _crlf ) ) );
            }
          }
          p.Close();
          sw.Close();
        }
      }
      return 0;
    }
  }
}

//3456789012345678901234567890123456789012345678901234567890
