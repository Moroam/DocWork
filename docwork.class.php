<?php
class DocWork{
  protected $phpWord = null;
  protected $section = null;

  const WRITERS = [
    'Word2007' => 'docx',
    'ODText' => 'odt',
    'RTF' => 'rtf'
  ];
  const CM = 566.92913385827; # twips in 1 CM
  const DEFAULT_FONT_SIZE = 12;
  const DEFAULT_FONT_NAME = 'Times New Roman';


  protected $INIpath = 'docwork.class.ini'; # same folder with docwork.class.php
  protected $defaultParagraphStyle = [
    'spaceAfter' => 0,
    'lineHeight' => 1
  ];
  protected $defaultFormat = 'Word2007';
  protected $defaultFontSize = self::DEFAULT_FONT_SIZE;
  protected $defaultFontName = self::DEFAULT_FONT_NAME;
  protected $pageSetting = [
    'marginLeft'   => 1*self::CM,
    'marginRight'  => 1*self::CM,
    'marginTop'    => 1*self::CM,
    'marginBottom' => 1*self::CM,
    'headerHeight' => 0.6*self::CM,
    'footerHeight' => 0.6*self::CM,
    'pageNumberingStart' => 1,
    'orientation' => 'portrait' # 'landscape'
  ];
  protected $headerFooterFontStyle = [
    'size' => self::DEFAULT_FONT_SIZE - 2,
    'color' => '111111'
  ];
  protected $headerFooterParagraphStyle = [
    'alignment'  => 'center'
  ];
  protected $templatePageNumbering = 'Страница {PAGE} из {NUMPAGES}.';
  protected $titleFontStyle = [
    'bold' => true
  ];
  protected $titleParagraphStyle = [
    'align' => 'center',
    'spaceBefore' => 160,
    'spaceAfter' => 160,
  ];
  protected $tabStyle = [
    'borderSize' => 0,
    'cellMargin' => 0,
    'alignment'  => 'center',
  ];
  protected $tabCellStyle = [
    'borderSize' => 0
  ];
  protected $tabHeadCellStyle = [
    'borderSize' => 0,
    'valign' => 'center',
    'align' => 'center',
  ];
  protected $tabFirstRowStyle = [
    'bgColor' => 'DDDDDD',
  ];


  public function __construct(array $pageSetting = [], string $ini_path = ''){
    $this->readINI($ini_path);
    $this->createDoc();
    $this->newSection($pageSetting);
  }

  /**
   * Read default INI values
   *
   */
  protected function readINI(string $ini_path = ''){
    if($ini_path != ''){
      if(file_exists($ini_path)){
        $this->INIpath = $ini_path;
      } else {
        throw new Exception("Error: ini file $ini_path don't exist", 1);
      }
    }

    $structure = parse_ini_file($this->INIpath, true, INI_SCANNER_NORMAL);

    # DEFAULT
    $this->defaultFormat         = $structure['DEFAULT']['format'   ]               ?? $this->defaultFormat;
    $this->defaultFontSize       = $structure['DEFAULT']['font_size']               ?? $this->defaultFontSize;
    $this->defaultFontName       = $structure['DEFAULT']['font_name']               ?? $this->defaultFontName;
    $this->templatePageNumbering = $structure['DEFAULT']['template_page_numbering'] ?? $this->templatePageNumbering;

    # PAGE SETTINGS
    $this->pageSetting               = array_merge($this->pageSetting,               $structure['PAGE SETTING']           ?? []);

    # STYLES
    $this->defaultParagraphStyle     = array_merge($this->defaultParagraphStyle,     $structure['DEFAULT PARAGRAPH']      ?? []);
    $this->headerFooterFontStyle     = array_merge($this->headerFooterFontStyle,     $structure['HEADER FOOTER FONT']     ?? []);
    $this->headerFooterParagraphStyle= array_merge($this->headerFooterParagraphStyle,$structure['HEADER FOOTER PARAGRAPH']?? []);
    $this->titleFontStyle            = array_merge($this->titleFontStyle,            $structure['TITLE FONT']             ?? []);
    $this->titleParagraphStyle       = array_merge($this->titleParagraphStyle,       $structure['TITLE PARAGRAPH']        ?? []);
    $this->tabStyle                  = array_merge($this->tabStyle,                  $structure['TABLE']                  ?? []);
    $this->tabCellStyle              = array_merge($this->tabCellStyle,              $structure['TABLE CELL']             ?? []);
    $this->tabHeadCellStyle          = array_merge($this->tabHeadCellStyle,          $structure['TABLE HEAD CELL']        ?? []);
    $this->tabFirstRowStyle          = array_merge($this->tabFirstRowStyle,          $structure['TABLE FIRST ROW']        ?? []);
  }

  /**
   * Create document, set default parameters, add table style, add title style
   *
   */
  protected function createDoc(){
    $this->phpWord = new \PhpOffice\PhpWord\PhpWord();
    $this->phpWord->setDefaultParagraphStyle($this->defaultParagraphStyle);
    $this->phpWord->setDefaultFontSize($this->defaultFontSize);
    $this->phpWord->setDefaultFontName($this->defaultFontName);

    $this->phpWord->getSettings()->setZoom(100);
    $this->phpWord->addTableStyle('Report Table no head', $this->tabStyle);
    $this->phpWord->addTableStyle('Report Table', $this->tabStyle, $this->tabFirstRowStyle);

    $this->phpWord->getSettings()->setThemeFontLang(new \PhpOffice\PhpWord\Style\Language(\PhpOffice\PhpWord\Style\Language::RU_RU));

    # Simple Title Styles. Depth => null, 1, 2, 3, 4
    $this->phpWord->addTitleStyle(
      null,
      array_merge($this->titleFontStyle, ['size' => $this->defaultFontSize + 8]),
      $this->titleParagraphStyle
    );
    $this->phpWord->addTitleStyle(
      1,
      array_merge($this->titleFontStyle, ['size' => $this->defaultFontSize + 6]),
      $this->titleParagraphStyle
    );
    $this->phpWord->addTitleStyle(
      2,
      array_merge($this->titleFontStyle, ['size' => $this->defaultFontSize + 4]),
      $this->titleParagraphStyle
    );
    $this->phpWord->addTitleStyle(
      3,
      array_merge($this->titleFontStyle, ['size' => $this->defaultFontSize + 2]),
      $this->titleParagraphStyle
    );
    $this->phpWord->addTitleStyle(
      4,
      array_merge($this->titleFontStyle, ['size' => $this->defaultFontSize]),
      $this->titleParagraphStyle
    );

    # Numbering Title Styles. Depth =>  10, 11, 12
    $this->phpWord->addNumberingStyle(
        'hNum',
        ['type' => 'multilevel', 'levels' => [
            ['pStyle' => 'Heading10', 'format' => 'decimal', 'text' => '%1.'      ],
            ['pStyle' => 'Heading11', 'format' => 'decimal', 'text' => '%1.%2.'   ],
            ['pStyle' => 'Heading12', 'format' => 'decimal', 'text' => '%1.%2.%3.'],
          ]
        ]
    );
    $this->phpWord->addTitleStyle(
      10,
      ['size' => $this->defaultFontSize + 4],
      ['numStyle' => 'hNum', 'numLevel' => 0]
    );
    $this->phpWord->addTitleStyle(
      11,
      ['size' => $this->defaultFontSize + 2],
      ['numStyle' => 'hNum', 'numLevel' => 1]
    );
    $this->phpWord->addTitleStyle(
      12,
      ['size' => $this->defaultFontSize],
      ['numStyle' => 'hNum', 'numLevel' => 2]
    );
  }

  /**
   * Add new section with, calc default table width (100% = page width - margin)
   */
  public function newSection(array $pageSetting = []){
    $this->section = $this->phpWord->addSection(array_merge($this->pageSetting, $pageSetting));
    $this->header = null;
    $this->footer = null;

    $sectionStyle = $this->section->getStyle();
    $tabWidth = $sectionStyle->getPageSizeW() - $sectionStyle->getMarginLeft() - $sectionStyle->getMarginRight();
    $this->tabStyle = array_merge(
      $this->tabStyle,
      ['width' => $tabWidth ] // table width allways 100%
    );
  }

  /**
   * Set Document information value by key
   */
  protected function setDocInfoByKey(string $key, string $value){
    $info_keys = [
      'creator' => 'Creator', 'company' => 'Company', 'title' => 'Title', 'category' => 'Category', 'lastmodifiedby' => 'LastModifiedBy',
      'description' => 'Description', 'created' => 'Created', 'modified' => 'Modified', 'subject' => 'Subject', 'keywords' => 'Keywords'
    ];
    $key = $info_keys[strtolower($key)] ?? '';
    if($key != ''){
      $doc_info = $this->phpWord->getDocInfo();
      $setKey = 'set' . $key;
      $doc_info->$setKey($value);
    }
  }

  /**
   * Set Document information
   */
  public function docInfo(array $properties = []){
    foreach ($properties as $key => $value) {
      $this->setDocInfoByKey($key, $value);
    }

    return $this;
  }

  /**
   * Add title to document
   * @param string $title
   * @param int $depth - depth / level of the title
   */
  public function title(string $title, int $depth = null){
    if($title!='') {
      $this->section->addTitle($title, $depth);
    }

    return $this;
  }

  /*
   * Add simple text to document
   */
  public function text(string $text, array $font_style = [], array $paragraph_style = []){
    if($text != '') {
      $this->section->addText($text, $font_style, $paragraph_style);
    }

    return $this;
  }

  /*
   * Add HTML to document
   */
  public function html(string $html){
    if ($html != ''){
      \PhpOffice\PhpWord\Shared\Html::addHtml($this->section, $html, false, false);
    }

    return $this;
  }

  /**
   * Add table to document
   *
   * @param array $data
   * @param array $options = [
   *    array  head     => values for first / head row
   *    array  width    => cells width
   *    string caption => table caption
   *    int    fontSize   => default font size for table
   *    int    tabWidth   => default = tabStyle['width'] (100% page width - margin)
   *    array  columns  => format for cells, may be set for individual cell column_number =>
   *             [
   *              'valign', 'bgColor', #### CELL STYLE
   *              'alignment',         #### PARAGRAPH STYLE
   *              'size', 'color', 'bold', 'italic', 'underline' ### FONT STYLE
   *             ]
   *  ]
   * Parameters value
   * 'valign' => Vertical alignment, top, center, both, bottom
   * 'alignment' => paragraph text alignment 'start', 'center', 'end', 'both', 'left', 'right', 'justify'
   * 'bold', 'italic' => true / false
   * 'underline' => 'none','dash','dashLong','dashLongHeavy','dotDash','dotDotDash','dotted','dottedHeavy','single','wavyHeavy','words'
   */
  public function table(array $data=[], array $options=[]){
    if(count($data) == 0){
     return $this;
    }

    // INIT OPTIONS
    $options['head'    ] = $options['head'    ] ?? [];
    $options['width'   ] = $options['width'   ] ?? [];
    $options['caption' ] = $options['caption' ] ?? '';
    $options['fontSize'] = $options['fontSize'] ?? $this->defaultFontSize - 2;
    $options['columns' ] = $options['columns' ] ?? [];
    $options['tabWidth'] = $options['tabWidth'] ?? $this->tabStyle['width'];

    $cols = count($data[0]);
    $options['width'] = $this->recalcWidth($options, $cols);
    $style = $this->calcStyle($options, $cols);

    $this->title($options['caption'], 3);

    if(count($options['head']) > 0){
      $table = $this->section->addTable('Report Table');
      $table->addRow();
      foreach ($options['head'] as $k => $val) {
        $table
          ->addCell($options['width'][$k], $this->tabHeadCellStyle)
          ->addText($val,
            ['size' => $options['fontSize'], 'bold' => true],
            ['alignment' => 'center']
          );
      }
    } else {
      $table = $this->section->addTable('Report Table no head');
    }

    $tags = fn($str) => strip_tags( str_replace('<br>', '. ', $str) ); #strip_tags <w:br />
    foreach ($data as $row){
      $table->addRow();
      foreach($row as $c => $val) {
        $table
          ->addCell($options['width'][$c], $style['cellStyle'][$c] )
          ->addText($tags($val), $style['fontStyle'][$c], $style['paragraphStyle'][$c]);
      }
    }

    return $this;
  }

  /**
   * Recalculate table cells width in proportion to the full width of the table
   */
  protected function recalcWidth(array $options, int $cols) : array {
    if(count($options['width'])==0){
      return array_fill(0, $cols, (int)($options['tabWidth'] / $cols));
    } else {
      for($arr = [], $s = 0, $i = 0; $i < $cols; $i++){
        $s += $arr[] = $options['width'][$i] ?? 0;
      }
      if($s == 0){
        return array_fill(0, $cols, (int)($options['tabWidth'] / $cols));
      }
      return array_map(fn($x)=>($x==0?1:(int)($options['tabWidth']*$x/$s)), $arr);
    }
  }

  /**
   * Calculate style for cells (cellStyle, paragraphStyle, fontStyle) from $options
   */
  protected function calcStyle(array $options, int $cols) : array {
    $cellStyle = [];
    $paragraphStyle = [];
    $fontStyle = [];

    $keys_cell_style = ['valign', 'bgColor'];
    $keys_paragraph_style = ['alignment'];
    $keys_font_style = ['size', 'color', 'bold', 'italic', 'underline'];

    $ma = fn($arr, $frmt, $keys) => array_merge($arr, array_filter($frmt, fn($k) => in_array($k, $keys), ARRAY_FILTER_USE_KEY) );

    $columns = $options['columns'];

    for($i = 0; $i < $cols; $i++){
      $cellStyle[] = $this->tabCellStyle;
      $paragraphStyle[] = [];
      $fontStyle[] = ['size' => $options['fontSize']];
      if(array_key_exists($i, $columns)){
        $cellStyle[$i]      = $ma($cellStyle[$i],      $columns[$i], $keys_cell_style     );
        $paragraphStyle[$i] = $ma($paragraphStyle[$i], $columns[$i], $keys_paragraph_style);
        $fontStyle[$i]      = $ma($fontStyle[$i],      $columns[$i], $keys_font_style     );
      }
    }

    return ['cellStyle' => $cellStyle, 'paragraphStyle' => $paragraphStyle, 'fontStyle' => $fontStyle];
  }

  /**
   * Return phpWord and section, else null
   */
  public function __get(string $name){
    if($name == 'section') return $this->section;
    if($name == 'phpWord') return $this->phpWord;

    return null;
  }

  public function textBreak(int $cnt = 1){
    $this->section->addTextBreak($cnt);
    return $this;
  }

  public function pageBreak(){
    $this->section->addPageBreak();
    return $this;
  }

  /**
   * Add header or footer
   */
  protected function addHeaderFooter(bool $header=true, $text='', array $font_style = [], array $paragraph_style = []){
    if($text != ''){
      $el = $header ? $this->section->addHeader() : $this->section->addFooter();
      $el->addPreserveText(
        $text,
        array_merge($this->headerFooterFontStyle, $font_style),
        array_merge($this->headerFooterParagraphStyle, $paragraph_style));
    }

    return $this;
  }

  /**
   * Add header
   */
  public function header(string $text = '', array $font_style = [], array $paragraph_style = []){
    return $this->addHeaderFooter(true, $text, $font_style, $paragraph_style);
  }

  /**
   * Add footer
   */
  public function footer(string $text = '', array $font_style = [], array $paragraph_style = []){
    return $this->addHeaderFooter(false, $text, $font_style, $paragraph_style);
  }


  /**
   * Add page number
   * @param string $template_page_numbering - by default 'Страница {PAGE} из {NUMPAGES}.'
   */
  public function pageNumber(string $place = 'header', array $font_style = [], array $paragraph_style = [], string $template_page_numbering = ''){
    $text = $template_page_numbering == '' ? $this->templatePageNumbering : $template_page_numbering;
    return $this->addHeaderFooter($place == 'header', $text, $font_style, $paragraph_style);
  }

  /**
   * Save document
   *
   * @param bool $local_save save document local file or send the file for download
   * @param string $file_path for local download
   */
  public function save(string $file_name = 'report', string $file_path = ''){
    if($file_path != ''){
      if(!is_dir($file_path)){
        echo "<h3>Не существует пути для сохранения файла: $file_path</h3>";
        return;
      }
      $this->phpWord->save($file_path . "/".$file_name.".".self::WRITERS[$this->defaultFormat], $this->defaultFormat);
      return;
    }

    $temp_file_uri = tempnam('', 'xyz');
    $this->phpWord->save($temp_file_uri, $this->defaultFormat);

    ob_start();
    if(ob_get_level()==0) ob_end_clean();
    ob_get_clean();
    ob_clean();
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="'.$file_name.'.'.self::WRITERS[$this->defaultFormat].'"');
    header('Content-Transfer-Encoding: binary');
    header('Expires: 0');
    header('Content-Length:'.filesize($temp_file_uri));
    readfile($temp_file_uri);
    unlink($temp_file_uri);
    exit;
  }
}
