from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle, KeepTogether
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


STUDENT_NAME = "Shourya Veer Singh"
ROLL_NUMBER = "________________"


sections = [
    ("Q1. Fundamental Principles of Data Compression",
     [
      ("Introduction", "Data compression is the process of representing information with fewer bits than the original representation while preserving the required level of information or quality. It is essential because digital data requires storage space, bandwidth, and transmission time. Compression reduces cost and improves communication speed, but it must maintain an acceptable trade-off between size, quality, and computational complexity."),
      ("Need for Data Compression", "Text, images, audio, and video commonly contain large volumes of repeated, correlated, or less important information. Compression supports faster file transfer, efficient use of network bandwidth, smaller backups, economical cloud storage, and quicker loading of multimedia on web and mobile devices."),
      ("Types of Redundancy", "Coding redundancy means that symbols are represented with more bits than necessary. Spatial redundancy occurs when neighbouring image pixels have similar values. Temporal redundancy occurs when consecutive frames in video are nearly identical. Inter-symbol redundancy arises when some symbol combinations occur more frequently or are predictable. Statistical redundancy exists when symbol probabilities are non-uniform."),
      ("Lossless and Lossy Compression", "Lossless compression guarantees exact reconstruction. It is used for text, executable files, source code, medical records, and archival data. Typical techniques include Run-Length Coding, Huffman Coding, Arithmetic Coding, LZW, and dictionary methods. Lossy compression discards information considered less perceptually important. It is used for photographs, audio, and video. JPEG, MPEG, MP3, and AAC are examples. The level of loss is controlled to achieve a desired quality–size trade-off."),
      ("Compression Ratio", "Compression ratio (CR) is usually expressed as C_R = original size/compressed size. It may also be expressed as percentage saving: 100 * (1 - compressed size/original size). For example, a 100 kB file stored in 25 kB has a compression ratio of 4:1 and saves 75% of the original storage."),
      ("Coding Efficiency", "Coding efficiency is the ratio of average information content (entropy) to the average code length: efficiency = H(S)/L_avg. It is close to 100% when the average code length approaches the entropy. For a symbol source, entropy is H(S) = -sum(p_i log2 p_i). It gives the theoretical lower bound on average bits per symbol for a uniquely decodable code."),
      ("Text-Compression Approaches", "Statistical coding assigns shorter codes to frequent symbols and longer codes to rare symbols. Huffman coding constructs a binary tree by repeatedly merging two least-frequent symbols. Arithmetic coding represents an entire message as a fraction in [0,1). Dictionary coding replaces recurring words or phrases with references to a dictionary; LZ77 uses a sliding window and LZ78/LZW build or transmit dictionary entries. Run-length coding is useful for long runs of identical symbols, but is not efficient for ordinary unstructured text without prior preprocessing.")
     ]),

    ("Q2. Static Huffman Coding",
     [
      ("Working Principle", "Static Huffman Coding uses symbol frequencies known before encoding. All symbols are placed in a priority queue ordered by frequency. The two least-frequent nodes are removed and joined as a parent whose weight is their sum. This process continues until one tree remains. The left branch is labelled 0 and the right branch is labelled 1, or vice versa, and the code for a symbol is obtained from the root-to-leaf path."),
      ("Construction", "Initial frequencies are E=10, C=15, D=15, B=20, and A=40. Merge E(10) and D(15) to obtain node X=25. Merge C(15) and B(20) to obtain Y=35. Merge X(25) and Y(35) to obtain Z=60. Finally merge Z=60 with A=40 to obtain root=100."),
      ("Tree", "Root(100)\n├── 0: A(40)\n└── 1: Z(60)\n    ├── 0: X(25)\n    │   ├── 0: E(10)\n    │   └── 1: D(15)\n    └── 1: Y(35)\n        ├── 0: C(15)\n        └── 1: B(20)\n\nTherefore, the left/right labelling above gives A=00, B=11, C=10, D=011, and E=010."),
      ("Average Code Length", "L_avg = (40/100)(2) + (20/100)(2) + (15/100)(2) + (15/100)(3) + (10/100)(3)\n= 0.80 + 0.40 + 0.30 + 0.45 + 0.30\n= 2.25 bits/symbol."),
      ("Total Bits", "The source contains 40 A, 20 B, 15 C, 15 D, and 10 E symbols. Total bits = 40(2) + 20(2) + 15(2) + 15(3) + 10(3) = 225 bits."),
      ("Entropy and Efficiency", "The entropy of this distribution is approximately 2.185 bits/symbol. Because 2.25 is close to 2.185, the coding is efficient. The exact codes may vary if equal-weight nodes are merged in a different valid order, but the average length should remain the optimum Huffman value.")
     ]),

    ("Q3. Dynamic Huffman Coding",
     [
      ("Concept", "Dynamic, or adaptive, Huffman Coding does not require all source frequencies beforehand. The encoder and decoder begin with an empty or NYT (Not-Yet-Transmitted) symbol set and update the Huffman tree after each symbol. Both sides perform the same update, so the tree remains synchronized."),
      ("Tree Construction and Update", "For the first occurrence of a symbol, the encoder sends the code for NYT and an identifier for the symbol. It then adds the symbol to the tree. For later occurrences, it sends the current code of the symbol. After every symbol, the node weight is increased and the tree is rearranged while satisfying Huffman/numbered-node invariants. A commonly used algorithm is FGK or Vitter."),
      ("Update Process", "The encoder finds the symbol leaf, increments its weight, and reorders nodes if necessary. Nodes are swapped with a highest-numbered node of the same weight only when permitted, preserving the sibling property. The decoder performs the same operations after receiving and interpreting each symbol."),
      ("Comparison", "Static Huffman coding stores the complete frequency table and initial code tree, and normally performs two passes: frequency collection and encoding. It needs little/no adaptive state at runtime but wastes no header? It requires a header or agreed statistics. Dynamic Huffman coding uses less pre-transmission model information and can adapt to changing symbol patterns, but requires more memory and processing for tree updates and is sensitive to synchronization errors."),
      ("Applications", "Static coding is suitable for archiving, document compression, and systems with known or stable distributions. Dynamic coding is useful for streaming or files whose symbol statistics are unknown or vary, especially in on-the-fly encoders.")
     ]),

    ("Q4. Arithmetic Coding",
     [
      ("Principle", "Arithmetic coding represents an entire message as a subinterval of [0,1). No individual symbol is assigned a fixed-width code. At each step, the current interval is divided in proportion to symbol probabilities. For symbols A, B, C with A=[0,0.5), B=[0.5,0.8), C=[0.8,1.0), the code sequence ABCA is processed as follows."),
      ("Interval Calculation", "Start: [0,1), width 1.\nA: [0,0.5), width 0.5.\nB: lower=0.5+0.5*0=0.5; upper=0.5+0.5*0.3=0.65, giving [0.5,0.65).\nC: lower=0.5+0.15*0.8=0.62; upper=0.5+0.15*1.0=0.65, giving [0.62,0.65).\nA: lower=0.62+0.03*0=0.62; upper=0.62+0.03*0.5=0.635, giving [0.62,0.635)."),
      ("Final Encoded Value", "Choose the midpoint of the final interval: (0.620 + 0.635)/2 = 0.6275. The value 0.6275 lies inside [0.62,0.635), so it uniquely identifies ABCA when the symbol probabilities and message length or termination convention are known."),
      ("Comparison with Huffman Coding", "Huffman coding assigns integer-length codes to symbols and is simple to implement. Arithmetic coding approaches the entropy closely even when symbol probabilities do not produce convenient fixed-length Huffman codes, and it can encode the whole sequence into a fractional value. Arithmetic coding is often more storage efficient for very skewed distributions, but it is numerically and computationally more complex. With long sequences, implementation uses finite-precision scaling and renormalization.")
     ]),

    ("Q5. Image Compression; GIF and TIFF",
     [
      ("Need and Basic Principles", "Images contain many pixels and may have spatial redundancy, limited colour variety, or content that the human eye cannot distinguish precisely. Compression reduces storage and transmission costs, speeds loading and printing, and enables archiving of large image collections."),
      ("GIF", "Features: GIF is a lossless, indexed-colour raster format using a palette of at most 256 colours. It supports interlacing, transparency, and short animation sequences. It commonly uses LZW dictionary compression. Advantages include wide browser support, small files for flat graphics, simple animation support, and no patent-issue concerns in modern use. Limitations include poor support for high-colour photographs, a fixed 8-bit palette per frame, limited control of compression quality, and larger files for complex real-world images."),
      ("TIFF", "Features: TIFF is a flexible, tagged, high-quality raster format. It can use RGB, greyscale, palette, CMYK, and many bit-depth configurations. It supports multiple pages/layers, metadata, scan data, and lossless methods such as LZW, PackBits, ZIP/Deflate, or none. Some TIFF variants support lossy JPEG compression. Advantages include excellent quality, high bit depth, broad professional and archival use, and strong metadata support. Limitations include very large file sizes, complexity due to many optional extensions, slower processing, and poor suitability for web delivery."),
      ("GIF versus TIFF", "GIF is mainly appropriate for logos, icons, simple diagrams, memes, and frame-by-frame animations. TIFF is used in document scanning, medical imaging, publishing, and print workflows where fidelity and metadata are important. TIFF provides higher quality and flexibility, while GIF prioritizes simple web-compatible graphics and small animated files.")
     ]),

    ("Q6. Digitized Documents",
     [
      ("Concept", "A digitized document is an electronic representation of a paper document. It may be a raster image, a searchable OCR (Optical Character Recognition) text layer, or a hybrid PDF. Digitization preserves accessibility and allows indexing, transmission, printing, editing, and backup."),
      ("Conversion Process", "The paper is prepared and scanned using a flatbed or sheet-fed scanner. Illumination and acquisition convert light reflected from the page into electrical signals. Sampling creates pixels, quantization assigns intensity levels, and optional colour correction or de-skewing improves the result. OCR then recognizes characters, and PDF or TIFF packages the image and searchable text."),
      ("Image Acquisition", "A scanner's optical system, sensor, and software determine the quality of acquisition. Clean optics, correct colour settings, flat pages, and controlled lighting reduce noise and shadows."),
      ("Resolution", "Resolution is measured in dots per inch (dpi). 150–300 dpi is typical for office documents, while 600 dpi or more may be used for fine print or archival scanning. Higher resolution captures more detail but increases storage and processing requirements."),
      ("Bit Depth", "Bit depth is the number of bits used for each pixel component. Binary 1-bit depth is suitable for clean black-and-white text, 8 bits per channel is common for greyscale or 256-level colour, and 24-bit colour has 8 bits for each RGB channel. More bit depth improves tonal fidelity but increases size."),
      ("Image Representation", "Text pages may be represented as bilevel or greyscale images. The page may be a bitmap or, after OCR, XML/text. PDF can contain a raster page with an invisible or searchable text layer."),
      ("Compression", "Lossless compression preserves exact pixel values and is used for financial, legal, archival, and OCR-critical documents. TIFF with LZW/ZIP or PDF with lossless image streams are common. Lossy JPEG or JBIG-type compression may be used only when controlled quality loss is acceptable. Compression reduces storage and transmission cost, enables faster indexing or download, and supports large document-management repositories.")
     ]),

    ("Q7. JPEG Image Compression",
     [
      ("Principle", "JPEG is a lossy compression standard for continuous-tone grayscale and colour images. It exploits the fact that natural images contain spatial redundancy and that human visual perception is less sensitive to high-frequency spatial detail."),
      ("Block Diagram", "Input Image -> RGB to YCbCr Conversion -> 8x8 Blocks -> Level Shifting -> 2-D DCT -> Quantization -> Zigzag Scan -> Run-Length Coding -> Entropy/Huffman Coding -> JPEG Bitstream"),
      ("Color Space Conversion", "The RGB image is transformed to a luminance component Y and chrominance components Cb and Cr. Since human vision is more sensitive to luminance, the chrominance planes are often subsampled 4:2:0 or 4:2:2."),
      ("DCT", "Each 8x8 block is transformed from the spatial domain to frequency coefficients using the Discrete Cosine Transform. One coefficient represents the average or DC value, and the remaining coefficients represent increasing spatial frequencies."),
      ("Quantization", "Each DCT coefficient is divided by a corresponding value in a quantization table and rounded. The larger table values for high frequencies remove fine detail and reduce precision, producing zeros and compression. The quantization table and quality factor control the quality–size trade-off."),
      ("Zigzag Scanning", "Coefficients are read in zigzag order from low frequency to high frequency. This tends to place non-zero coefficients first and long runs of zeros later."),
      ("Run-Length Encoding", "Zero coefficients are encoded using run-length pairs (run of zeros, non-zero value). End-of-Block markers indicate that the remaining coefficients in a block are zero."),
      ("Huffman Coding", "Final values and run/amplitude symbols are represented by variable-length Huffman codes. Frequent symbols receive short codes, and rare symbols receive longer codes."),
      ("Lossy versus Lossless", "JPEG baseline mode is lossy. Loss occurs primarily at quantization and cannot be fully recovered. Lossless JPEG modes and formats such as PNG or TIFF with lossless coding allow exact reconstruction of the original pixels.")
     ]),

    ("Q8. Differential Pulse Code Modulation (DPCM)",
     [
      ("Principle", "Pulse Code Modulation encodes each sample independently. DPCM reduces redundancy by encoding only the difference between the actual sample and a prediction of that sample. Adjacent samples in a slowly varying signal are often similar, so the prediction error has a smaller range than the original sample."),
      ("Block Diagram", "Input x[n] -> Difference Unit (+/-) -> Quantizer -> Encoder/Transmitter\n                                           ^\n                                           |\n                Predictor <- Quantized Difference / Reconstructed Value"),
      ("Predictor", "The predictor estimates the next sample from previous reconstructed samples. A simple predictor may output x_hat[n] = x[n-1]. In practice, a linear predictor uses several past samples."),
      ("Quantizer", "The quantizer maps the prediction error e[n] = x[n] - x_hat[n] to a finite set of levels. This introduces quantization error and determines the output bit rate."),
      ("Encoder", "The encoder converts the quantizer index into a binary word. The receiver uses the same predictor and adds the decoded difference to the predicted value to reconstruct the signal."),
      ("Advantages", "DPCM uses fewer bits per sample than PCM for correlated signals, provides lower bit rate for similar quality, and can be implemented relatively easily."),
      ("Limitations", "It is sensitive to prediction and channel errors, requires feedback or the same reconstructed history at transmitter and receiver, and may suffer poor performance for signals that are difficult to predict.")
     ]),

    ("Q9. Adaptive Differential Pulse Code Modulation (ADPCM)",
     [
      ("Concept", "ADPCM improves DPCM by adapting the quantizer step size and/or predictor to signal characteristics. It therefore works well for both small- and large-amplitude variations."),
      ("Block Diagram", "Input x[n] -> Difference -> Adaptive Quantizer -> Encoder -> Bitstream\n                                |                    |\n                         Step-Size Adaptation <- Quantizer Index\n                                |\n                         Adaptive Predictor <- Reconstructed x[n]"),
      ("Adaptation of Step Size", "When successive quantizer indices indicate that the signal is highly varying or the error is large, the step-size adaptation logic increases Delta. If the signal is slowly varying and errors are small, Delta is decreased. Multiplication factors, look-up tables, or recursive adaptation rules are used."),
      ("Working", "The current prediction is subtracted from the input. The adaptive quantizer encodes the resulting difference using a dynamically controlled step size. The decoder updates the same step size and predictor to reconstruct the sample."),
      ("PCM, DPCM, and ADPCM Comparison", "PCM encodes each sample independently and has high predictability but a high bit rate. DPCM encodes prediction error and is efficient for correlated signals. ADPCM additionally adapts to signal amplitude and is most efficient in variable speech and audio applications. ADPCM usually has higher complexity than DPCM and PCM, but it gives lower bit rates with better perceived quality.")
     ]),

    ("Q10. Adaptive Subband Coding",
     [
      ("Concept", "Subband coding divides an audio signal into several frequency bands, typically using analysis filter banks. Each band is coded independently according to its signal power and perceptual importance."),
      ("Division into Subbands", "A low-pass/high-pass quadrature mirror filter bank or a related filter bank separates the input. Repeated filtering and downsampling create low, mid, and high-frequency bands. The bands are approximately orthogonal or can be reconstructed by a synthesis filter bank."),
      ("Independent Encoding", "Each subband is sampled at a rate appropriate to its bandwidth and may use PCM or ADPCM coding. High-energy bands and perceptually important bands receive more bits."),
      ("Adaptive Bit Allocation", "The encoder estimates the energy, spectral content, and masking level of each band. It allocates bits according to a distortion criterion and perceptual weighting. Bands with little power or already masked content receive fewer or zero bits."),
      ("Advantages", "Subband coding separates unlike frequency characteristics, reduces quantization noise in low-energy bands, supports efficient adaptive allocation, and can be combined with psychoacoustic masking. It offers better quality than a simple wideband coder at the same bit rate."),
      ("Applications", "It is used in speech and audio compression, including MPEG audio and digital broadcasting, especially where quality and low bit rate are important.")
     ]),

    ("Q11. Delta Modulation",
     [
      ("Principle", "Delta Modulation represents a signal using only 1 bit per sample. The transmitter compares the input sample with a predicted/reconstructed sample and sends 1 if the input is greater and 0 if it is smaller. The step size Delta determines the reconstruction change."),
      ("Block Diagram", "Input x[n] -> Comparator -> 1-bit Quantizer/Encoder -> Bitstream\n                  ^                  |\n                  |                  v\n               Local Predictor/Accumulator <- Bit"),
      ("Working", "The accumulator adds +Delta or -Delta according to the received bit. The resulting staircase waveform tracks a slowly changing input. The quantization error is the difference between the input and staircase approximation."),
      ("Slope Overload Distortion", "If the input changes faster than the accumulator can follow, the waveform accumulates only +Delta or -Delta repeatedly and falls behind. This produces slope overload distortion."),
      ("Granular Noise", "If Delta is much larger than the small variations of a slowly varying signal, the staircase oscillates around the input. This produces granular noise."),
      ("Adaptive Delta Modulation", "ADM changes Delta according to the bit pattern or signal slope. If several consecutive bits have the same direction, Delta increases to follow rapid changes and reduce slope overload. If the direction alternates or changes indicate a slow segment, Delta decreases to reduce granular noise."),
      ("DM versus ADM", "DM uses a fixed step size, is simple, and has 1-bit encoding, but is sensitive to both overload and granular noise. ADM adapts the step size, usually provides better signal-to-noise ratio and can cover more input ranges, but requires more logic and synchronized step-size adaptation.")
     ]),

    ("Q12. Low-Bit-Rate Speech Coding and Vocoders",
     [
      ("Challenges", "Speech varies rapidly in time and frequency and is strongly speaker dependent. Low bit rates must remove not only waveform redundancy but also predictable spectral structure. The coder must preserve intelligibility and naturalness, handle background noise, avoid pitch and timing errors, and remain robust to packet loss."),
      ("Vocoder Principle", "A vocoder models speech production instead of waveform values. It typically separates an excitation source (voiced pulse train or unvoiced noise) from a spectral/vocal-tract filter. Parameters are extracted and transmitted at frame intervals."),
      ("Block Diagram", "Analysis side:\nSpeech -> Frame Windowing -> Pitch/Voicing Analysis -> Spectral Envelope/LPC Analysis -> Parameter Encoder -> Bitstream\n\nSynthesis side:\nBitstream -> Parameter Decoder -> Excitation Generator -> Vocal-Tract Synthesis Filter -> Reconstructed Speech"),
      ("Operation", "The analysis stage divides speech into frames, estimates pitch or voicing, computes filter coefficients or an equivalent spectral representation, and estimates gain. The synthesizer recreates voiced excitation or noise and passes it through the vocal-tract filter."),
      ("Applications", "Low-bit-rate vocoders are used in mobile telephony, secure communications, satellite links, military radio, voice storage, and bandwidth-limited networks."),
      ("Limitations", "They may produce robotic or buzzy speech, can lose fine speaker identity, perform poorly in noisy environments, and can be sensitive to frame loss. Naturalness is often lower than waveform coders.")
     ]),

    ("Q13. Linear Predictive Coding and CELP",
     [
      ("LPC Principle", "LPC assumes that each speech sample can be approximated by a linear combination of previous samples. The vocal tract is represented by an all-pole filter. LPC analysis estimates coefficients that minimize the prediction residual."),
      ("LPC Analysis", "Speech is divided into overlapping frames and windowed. An autocorrelation or covariance method computes predictor coefficients. A frame gain and pitch/voicing parameters may also be estimated. The residual is the error or excitation that drives the vocal-tract model."),
      ("LPC Synthesis", "At the decoder, the excitation source is generated by a periodic pulse train for voiced speech or noise for unvoiced speech. It is filtered by the LPC synthesis filter to produce speech."),
      ("Block Diagram", "Analysis:\nSpeech -> Window -> LPC Coefficient Analysis -> Residual/Voicing/Pitch/Gain\n\nSynthesis:\nExcitation Source -> LPC Synthesis Filter -> Speech"),
      ("CELP", "Code Excited Linear Prediction uses LPC synthesis filters but improves excitation through a codebook of candidate excitation vectors. It performs an analysis-by-synthesis search: candidate vectors are filtered and compared with the original speech, and the best codebook index and gain are transmitted."),
      ("CELP Improvement", "CELP improves naturalness and reduces buzziness because it represents both periodic and noise-like details through adaptive and fixed codebook excitation. It can capture a wider range of speech events than simple pitch/noise excitation."),
      ("LPC versus CELP", "Traditional LPC coding is less complex and has lower delay, but uses approximate excitation and can sound less natural. CELP provides substantially better quality and lower perceived distortion, but requires codebook search, more computation, more storage, and sometimes higher bit rate.")
     ]),

    ("Q14. Perceptual Audio Coding",
     [
      ("Principle", "Perceptual coding removes or reduces components that are inaudible or less noticeable to a human listener. The encoder uses psychoacoustic models instead of treating all quantization errors equally."),
      ("Frequency Masking", "A strong tone or narrow-band sound makes neighbouring frequencies inaudible. A masker can hide weaker spectral components located close in frequency. The encoder allocates no or very few bits to frequency bins below the masking threshold."),
      ("Temporal Masking", "A strong sound can mask weaker sounds immediately before and after it. Pre-masking is short, while post-masking lasts longer. Temporal masking allows selected samples or transform coefficients to be coded with higher noise."),
      ("MPEG Audio", "MPEG audio uses subband or hybrid transform coding, psychoacoustic masking, MDCT-like transforms in later layers, block selection, and Huffman coding. MPEG Layer I, II, and III (MP3) provide increasing efficiency and complexity. MP3 separates frequencies, allocates bits according to masking, and uses Huffman entropy coding."),
      ("Dolby Audio", "Dolby audio systems use perceptual coding, subband or transform filtering, psychoacoustic masking, and proprietary or standardized metadata and processing. Dolby Digital (AC-3) and modern Dolby formats may include multichannel coding and dialog, dynamic-range, or surround metadata."),
      ("Comparison", "MPEG audio is an international, widely interoperable family of perceptual standards used in media, streaming, broadcasting, and digital video. Dolby systems are often optimized for consumer and cinema multichannel experiences and may use proprietary metadata and delivery tools. Both rely on masking and scale-factor/bit-allocation techniques. MPEG MP3 is widely used for general stereo music, while Dolby systems are prominent in film, television, and surround sound.")
     ]),

    ("Q15. Video Compression; H.261 and MPEG",
     [
      ("Fundamental Principles", "Video contains spatial redundancy within frames and temporal redundancy between frames. The codec exploits repeated textures, similar neighboring pixels, moving objects, and stationary backgrounds to reduce representation size."),
      ("Spatial and Temporal Redundancy", "Spatial redundancy arises because neighbouring pixels and macroblocks are often similar. Transform coding, quantization, and entropy coding exploit it. Temporal redundancy means consecutive frames contain similar regions. Motion estimation and compensation exploit this redundancy."),
      ("Motion Estimation", "The encoder compares macroblocks or regions in a current frame with candidate regions in a reference frame. A motion vector indicates the best displacement. Metrics such as mean absolute difference or sum of absolute difference are often used."),
      ("Motion Compensation", "The predicted block is reconstructed by shifting the matching block from the reference frame using the motion vector. The encoder stores motion vectors and the residual difference rather than all original pixels."),
      ("Frame Prediction", "Intra-frame prediction encodes a frame independently. Inter-frame prediction uses one or more previously decoded or future frames to estimate the current frame."),
      ("I-Frames", "I-frames are intra-coded frames. They do not use motion prediction from other frames and provide random-access points. They usually offer the highest size and best decode independence."),
      ("P-Frames", "P-frames are predicted from a previous reference frame. They store motion vectors and residuals and are smaller than I-frames when prediction is accurate."),
      ("B-Frames", "B-frames use both previous and future reference frames for bidirectional prediction. They can achieve greater compression but require frame reordering and additional delay."),
      ("H.261", "H.261 was designed for videoconferencing at approximately p x 64 kbit/s. It uses CIF/QCIF video, 8x8 DCT blocks, macroblocks, motion compensation, quantization, run-length coding, and variable-length coding. It mainly targets low delay and supports I- and P-frames."),
      ("MPEG", "MPEG standards target stored and broadcast digital video. MPEG-1 was used for CD-quality video, MPEG-2 for DVDs and television broadcasting, and MPEG-4/H.264/HEVC and later variants support more resolutions, coding tools, and lower bit rates. MPEG introduces I-, P-, and B-frame coding, scalable GOP structures, bidirectional prediction, and sophisticated motion compensation."),
      ("H.261 versus MPEG", "H.261 emphasizes low-delay conversational video and simpler prediction, while MPEG emphasizes storage or broadcast efficiency and may use bidirectional B-frames, flexible GOP structures, and higher compression. MPEG systems generally provide better compression and quality for a given application, but can introduce delay and decoder complexity.")
     ])
]


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def add_doc_header_footer(section):
    header = section.header
    table = header.add_table(rows=1, cols=2, width=Inches(6.5))
    table.columns[0].width = Inches(4.0)
    table.columns[1].width = Inches(2.5)
    table.cell(0, 0).text = "Name: " + STUDENT_NAME
    table.cell(0, 1).text = "Roll No.: " + ROLL_NUMBER + "\nSignature: __________________"
    for cell in table.rows[0].cells:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(9)
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("ITC Assignment 01  |  ")
    run.font.size = Pt(9)
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.append(fld_char1)
    run._r.append(instr_text)
    run._r.append(fld_char2)


def build_docx(path):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.65)
    sec.bottom_margin = Inches(0.65)
    sec.left_margin = Inches(0.75)
    sec.right_margin = Inches(0.75)
    add_doc_header_footer(sec)

    styles = doc.styles
    styles["Normal"].font.name = "Times New Roman"
    styles["Normal"].font.size = Pt(11)
    styles["Title"].font.name = "Times New Roman"

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("NATIONAL INSTITUTE OF TECHNOLOGY RAIPUR")
    r.bold = True
    r.font.size = Pt(16)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Department of Information Technology")
    r.bold = True
    r.font.size = Pt(13)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ASSIGNMENT 01")
    r.bold = True
    r.font.size = Pt(15)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("Course: ITC  |  B.Tech. - Fifth Semester").italic = True
    doc.add_paragraph("")

    info = doc.add_table(rows=3, cols=2)
    info.alignment = WD_TABLE_ALIGNMENT.CENTER
    info.style = "Table Grid"
    values = [
        ("Name", STUDENT_NAME),
        ("Roll Number", ROLL_NUMBER),
        ("Signature", "________________"),
    ]
    for i, (label, value) in enumerate(values):
        info.cell(i, 0).text = label
        info.cell(i, 1).text = value
        set_cell_shading(info.cell(i, 0), "D9E2F3")

    doc.add_page_break()

    for title, parts in sections:
        p = doc.add_paragraph()
        p.style = doc.styles["Heading 1"]
        p.add_run(title)
        for heading, text in parts:
            hp = doc.add_paragraph()
            hr = hp.add_run(heading)
            hr.bold = True
            hr.font.color.rgb = RGBColor(31, 78, 121)
            for line in text.split("\n"):
                doc.add_paragraph(line)
        doc.add_page_break()

    doc.save(path)


try:
    pdfmetrics.registerFont(TTFont("TimesNewRoman", r"C:\Windows\Fonts\times.ttf"))
    pdfmetrics.registerFont(TTFont("TimesNewRoman-Bold", r"C:\Windows\Fonts\timesbd.ttf"))
    font_name = "TimesNewRoman"
    bold_name = "TimesNewRoman-Bold"
except Exception:
    font_name = "Helvetica"
    bold_name = "Helvetica-Bold"


def pdf_header_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont(font_name, 9)
    canvas.drawString(2 * cm, A4[1] - 1.3 * cm, "Name: " + STUDENT_NAME)
    canvas.drawRightString(A4[0] - 2 * cm, A4[1] - 1.3 * cm, "Roll No.: " + ROLL_NUMBER)
    canvas.drawRightString(A4[0] - 2 * cm, A4[1] - 0.9 * cm, "Signature: __________________")
    canvas.setFont(font_name, 8)
    canvas.drawCentredString(A4[0] / 2, 1.1 * cm, "ITC Assignment 01")
    canvas.drawRightString(A4[0] - 2 * cm, 1.1 * cm, str(canvas.getPageNumber()))
    canvas.restoreState()


def build_pdf(path):
    styles = getSampleStyleSheet()
    normal = ParagraphStyle(
        "AssignmentNormal", parent=styles["Normal"],
        fontName=font_name, fontSize=10.5, leading=14,
        textColor=colors.HexColor("#222222"), spaceAfter=7,
    )
    heading1 = ParagraphStyle(
        "AssignmentHeading1", parent=styles["Heading1"],
        fontName=bold_name, fontSize=14, leading=17,
        textColor=colors.HexColor("#1F4E79"), spaceBefore=8, spaceAfter=10,
        keepWithNext=True,
    )
    subheading = ParagraphStyle(
        "AssignmentSubheading", parent=styles["Heading2"],
        fontName=bold_name, fontSize=11.5, leading=14,
        textColor=colors.HexColor("#1F4E79"), spaceBefore=7, spaceAfter=4,
        keepWithNext=True,
    )
    title = ParagraphStyle(
        "AssignmentTitle", parent=styles["Title"],
        fontName=bold_name, fontSize=16, leading=20,
        alignment=TA_CENTER, textColor=colors.HexColor("#1F1F1F"),
    )

    doc = SimpleDocTemplate(
        path, pagesize=A4,
        rightMargin=2 * cm, leftMargin=2 * cm,
        topMargin=2.1 * cm, bottomMargin=1.7 * cm,
    )
    story = []
    story.append(Paragraph("NATIONAL INSTITUTE OF TECHNOLOGY RAIPUR", title))
    story.append(Spacer(1, 0.25 * cm))
    story.append(Paragraph("Department of Information Technology", ParagraphStyle(
        "dept", parent=title, fontSize=13, leading=16)))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph("ASSIGNMENT 01", ParagraphStyle(
        "assignment", parent=title, fontSize=15, leading=18)))
    story.append(Spacer(1, 0.15 * cm))
    story.append(Paragraph("Course: ITC  |  B.Tech. - Fifth Semester", ParagraphStyle(
        "course", parent=normal, alignment=TA_CENTER)))
    story.append(Spacer(1, 0.5 * cm))
    info = [
        ["Name", STUDENT_NAME],
        ["Roll Number", ROLL_NUMBER],
        ["Signature", "________________"],
    ]
    table = Table(info, colWidths=[4 * cm, 10 * cm])
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#D9E2F3")),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(table)
    story.append(PageBreak())

    for index, (title_text, parts) in enumerate(sections):
        story.append(Paragraph(title_text, heading1))
        for heading, text in parts:
            story.append(Paragraph(heading, subheading))
            for line in text.split("\n"):
                story.append(Paragraph(line.replace("&", "&amp;"), normal))
        if index != len(sections) - 1:
            story.append(PageBreak())

    doc.build(story, onFirstPage=pdf_header_footer, onLaterPages=pdf_header_footer)


if __name__ == "__main__":
    docx_path = r"D:\daa pyq\DAA-PYQS\Shourya Veer Singh - ITC Assignment 01.docx"
    pdf_path = r"D:\daa pyq\DAA-PYQS\Shourya Veer Singh - ITC Assignment 01.pdf"
    build_docx(docx_path)
    build_pdf(pdf_path)
    print(docx_path)
    print(pdf_path)
