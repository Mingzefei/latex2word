import os
import glob
import subprocess
import concurrent.futures
import shutil
import logging
import regex
from tqdm import tqdm
import uuid

# Templates and patterns
MULTIFIG_TEXFILE_TEMPLATE = r"""
\documentclass[preview,convert,convert={outext=.png,command=\unexpanded{pdftocairo -r 600 -png \infile}}]{standalone}
\usepackage{graphicx}
\usepackage{subfig}
\graphicspath{{%s}}

\begin{document}
\thispagestyle{empty}
%s
\end{document}
"""
MULTIFIG_FIGURE_TEMPLATE = r"""
\begin{figure}[htbp]
    \centering
    \includegraphics[width=0.8\linewidth]{%s}
    \caption{%s}
    \label{%s}
\end{figure}
"""
FIGURE_PATTERN = r'\\begin{figure}.*?\\end{figure}'
CAPTION_PATTERN = r'\\caption(\{(?:[^{}]|(?1))*\})'
LABEL_PATTERN = r'\\label\{(.*?)\}'
GRAPHICSPATH_PATTERN = r'\\graphicspath\{\{(.+?)\}\}'

class LatexToWordConverter:
    def __init__(self, input_texfile, multifig_dir, output_docxfile, reference_docfile,bibfile,cslfile,debug=False):
        """
        Initializes the main class of the latex2word tool.

        Args:
            input_texfile (str): The path to the input LaTeX file.
            multifig_dir (str): The directory where the generated multi-figure LaTeX files will be stored.
            output_docxfile (str): The path to the output Word document file.
            reference_docfile (str): The path to the reference Word document file.
            bibfile (str): The path to the BibTeX file.
            cslfile (str): The path to the CSL file.
            debug (bool, optional): Whether to enable debug mode. Defaults to False.
        """
        # Initialize file paths
        self.input_texfile = os.path.abspath(input_texfile)
        self.multifig_dir = os.path.abspath(multifig_dir)
        self.output_texfile = os.path.abspath(input_texfile.replace('.tex', '_modified.tex'))
        self.output_docxfile = os.path.abspath(output_docxfile)
        self.reference_docfile = os.path.abspath(reference_docfile)
        self.bibfile = os.path.abspath(bibfile)
        self.cslfile = os.path.abspath(cslfile)
        self.luafile = os.path.join(os.path.dirname(os.path.abspath(__file__)),'resolve_equation_labels.lua')

        # Initialize other attributes
        self._raw_content = None
        self._modified_content = None
        self._raw_fig_contents = None
        self._raw_graphicspath = None
        self.generated_multifig_texfiles = set()

        # Initialize logger
        self.logger = logging.getLogger(f"{__name__}_{uuid.uuid4().hex[:6]}")
        handler = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)
        self.logger.setLevel(logging.DEBUG if debug else logging.INFO)
        
        self.logger.debug(f'Init input texfile: {self.input_texfile}')
        self.logger.debug(f'Init multifig_dir: {self.multifig_dir}')
        self.logger.debug(f'Init output texfile: {self.output_texfile}')
        self.logger.debug(f'Init output docxfile: {self.output_docxfile}')
        self.logger.debug(f'Init reference docfile: {self.reference_docfile}')
        self.logger.debug(f'Init bibfile: {self.bibfile}')
        self.logger.debug(f'Init cslfile: {self.cslfile}')
        self.logger.debug(f'Init luafile: {self.luafile}')

    def _match_pattern(self, pattern, content, multiple=False):
        """
        Matches a pattern in the given content and returns the result.

        Args:
            pattern (str): The regular expression pattern to match.
            content (str): The content to search for the pattern.
            multiple (bool, optional): Specifies whether to find multiple matches or not. Defaults to False.

        Returns:
            str or list: The matched result(s) if found, otherwise None.

        """
        if multiple:
            return regex.findall(pattern, content, regex.DOTALL)
        else:
            match = regex.search(pattern, content)
            return match.group(1) if match else None

    def analyze_texfile(self):
        """
        Analyzes the input LaTeX file and extracts relevant information.

        This method reads the content of the input LaTeX file, matches patterns to extract figure environments,
        and determines the path to the directory containing the figures.

        Returns:
            None
        """
        with open(self.input_texfile, 'r') as file:
            self._raw_content = file.read()
        self._raw_fig_contents = self._match_pattern(FIGURE_PATTERN, self._raw_content, multiple=True)

        graphicspath = self._match_pattern(GRAPHICSPATH_PATTERN, self._raw_content)
        if graphicspath:
            self._raw_graphicspath = os.path.abspath(os.path.join(os.path.dirname(self.input_texfile), graphicspath))
        else:
            self._raw_graphicspath = os.path.abspath(os.path.dirname(self.input_texfile))

            self.logger.debug(f'Init input figure directory(_raw_graphicspath): {self._raw_graphicspath}')
        
        self.logger.info(f'Read {os.path.basename(self.input_texfile)} texfile and found {len(self._raw_fig_contents)} figenvs.')

    def generate_multifig_texfiles(self):
        """
        Generates multiple tex files for each figure content.

        This method takes the raw figure contents and generates a tex file for each figure content.
        It comments out the captions in LaTeX and uses labels or default counters as filenames.

        Returns:
            None
        """
        default_counter = 1

        if os.path.exists(self.multifig_dir):
            shutil.rmtree(self.multifig_dir)
        os.makedirs(self.multifig_dir)

        for figure_content in self._raw_fig_contents:
            # Define a function to prepend a '%' character to each caption line
            # This effectively comments out the caption lines in LaTeX
            def comment_out_caption(match):
                # Add a '%' character before each caption line
                # Also add a '%' character before each newline character within the caption
                return '% \\caption' + match.group(1).replace('\n', '\n%')

            # Apply the function to each caption in the figure content
            # This comments out all captions in the figure content
            processed_figure_content = regex.sub(CAPTION_PATTERN, comment_out_caption, figure_content)

            # Find the last label of the figure
            labels = self._match_pattern(LABEL_PATTERN, processed_figure_content, multiple=True)
            if labels:
                # If a label is found, use the last one as the filename
                filename = labels[-1]
                # Remove the 'fig:' prefix if it exists
                if filename.startswith('fig:'):
                    filename = filename[4:]
            else:
                # If no label is found, use the default counter as the filename
                filename = f'fig{default_counter}'
                default_counter += 1

            # Check if the filename is already used
            while filename in self.generated_multifig_texfiles:
                filename = f'{filename}_{default_counter}'
                default_counter += 1

            # Add the filename to the set of generated filenames
            self.generated_multifig_texfiles.add(filename)
            
            file_content = MULTIFIG_TEXFILE_TEMPLATE % (os.path.abspath(os.path.join('..', self._raw_graphicspath)), processed_figure_content)

            # Create the tex file
            file_path = os.path.join(self.multifig_dir, f'{filename}.tex')
            with open(file_path, 'w') as file:
                file.write(file_content)

            self.logger.info(f'Created texfile {os.path.basename(file_path)} under {os.path.basename(self.multifig_dir)}.')

    def compile_multifig_texfile(self, texfile):
        """
        Compiles a LaTeX file containing multiple figures.

        Args:
            texfile (str): The path to the LaTeX file to be compiled.

        Returns:
            str: The path to the compiled LaTeX file.

        Raises:
            subprocess.CalledProcessError: If the compilation fails.

        """
        if self.logger.level == logging.DEBUG:
            subprocess.run(['xelatex', '-shell-escape', '-synctex=1' ,'-interaction=nonstopmode', texfile], check=True)
        else:
            subprocess.run(['xelatex', '-shell-escape', '-synctex=1' ,'-interaction=nonstopmode', texfile], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return texfile

    def compile_multifig_texfiles(self):
        """
        Compiles multiple .tex files and generates corresponding .png files.
        
        This method changes the current working directory to the output directory specified by `self.multifig_dir`.
        It then uses a `ProcessPoolExecutor` to compile each .tex file in parallel. After each .tex file is compiled,
        it checks for the generated .png files and performs the necessary checks and renaming operations.
        Finally, it changes back to the original working directory.
        
        Returns:
            None
        """
        # Save the current working directory
        cwd = os.getcwd()

        # Change to the output directory
        os.chdir(self.multifig_dir)

        try:
            # Compile each .tex file
            with concurrent.futures.ProcessPoolExecutor() as executor:
                futures = {executor.submit(self.compile_multifig_texfile, texfile) for texfile in self.generated_multifig_texfiles}
                for future in tqdm(concurrent.futures.as_completed(futures), total=len(futures), desc='Compiling texfiles'):
                    # print(f'Compiled tex file {future.result()}.tex')

                    # Check and Rename
                    generated_pngfiles = glob.glob(f'{future.result()}-*.png')
                    num_files = len(generated_pngfiles)

                    if num_files == 0:
                        self.logger.info(f'Error: No pngfile generated for {future.result()}')
                    elif num_files > 1:
                        self.logger.info(f'Warning: Multiple pngfiles generated for {future.result()}')

                    if num_files >= 1:
                        new_name = future.result() + '.png'
                        os.rename(generated_pngfiles[0], new_name)
        finally:
            # Change back to the original working directory
            os.chdir(cwd)

        self.logger.info('Created and renamed pngfiles.')

    def generate_modified_texfile(self):
        """
        Generates a modified .tex file by replacing figure contents and updating \graphicspath.

        This method iterates over the figure contents in the raw content of the .tex file and replaces them with modified
        figure contents. It also updates the \graphicspath to the specified multifig_dir. Finally, it writes the modified
        content to a new .tex file.

        Returns:
            None
        """
        # Assume png_files is a list of generated png files
        png_files = sorted(glob.glob(os.path.join(self.multifig_dir, '*.png')))

        # For each figure_content in figure_contents
        for i, figure_content in enumerate(self._raw_fig_contents):
            # Define the new figure content
            caption = self._match_pattern(CAPTION_PATTERN, figure_content)
            label = self._match_pattern(LABEL_PATTERN, figure_content)
            png_file = os.path.basename(png_files[i])
            modified_figure_content = MULTIFIG_FIGURE_TEMPLATE % (png_file, caption, label)
            
            # Replace the old figure content with the new figure content
            self._modified_content = self._raw_content.replace(figure_content, modified_figure_content)

        # Redefine \graphicspath
        self._modified_content = regex.sub(GRAPHICSPATH_PATTERN, r'\\graphicspath{{%s}}' % self.multifig_dir, self._modified_content)

        # Write the modified text content to a new .tex file
        with open(self.output_texfile, 'w') as f:
            f.write(self._modified_content)

        self.logger.info(f'Created {os.path.basename(self.output_texfile)} tex file.')

    def convert_modified_texfile(self):
        """
        Converts a modified TeX file to a DOCX file using pandoc.

        Raises:
            Exception: If pandoc or pandoc-crossref is not installed.

        Returns:
            None
        """
        # Check if pandoc and pandoc-crossref are installed
        if shutil.which('pandoc') is None:
            raise Exception('pandoc is not installed. Please install it before running this script.')
        if shutil.which('pandoc-crossref') is None:
            raise Exception('pandoc-crossref is not installed. Please install it before running this script.')

        # Define the command
        command = [
            'pandoc', self.output_texfile, '-o', self.output_docxfile,
            '--lua-filter', self.luafile,
            '--filter', 'pandoc-crossref',
            '--reference-doc='+self.reference_docfile,
            '--number-sections',
            '-M', 'autoEqnLabels', 
            '-M', 'tableEqns',
            '-M', 'reference-section-title=Reference',
            '--bibliography='+self.bibfile,
            '--citeproc', '--csl', self.cslfile
        ]

        # Save the current working directory
        cwd = os.getcwd()
        # Change to the output directory
        os.chdir(os.path.dirname(self.output_texfile))

        try:
            # Run the command
            subprocess.run(command, check=True)
        finally:
            # Change back to the original working directory
            os.chdir(cwd)

        self.logger.info(f'Converted {os.path.basename(self.output_texfile)} texfile to {os.path.basename(self.output_docxfile)} docxfile.')

    def convert(self):
        """
        Converts a LaTeX file to Word format.

        This method performs the following steps:
        1. Analyzes the LaTeX file.
        2. Generates multiple figure LaTeX files.
        3. Compiles the multiple figure LaTeX files.
        4. Generates the modified LaTeX file.
        5. Converts the modified LaTeX file to Word format.
        """
        self.analyze_texfile()
        self.generate_multifig_texfiles()
        self.compile_multifig_texfiles()
        self.generate_modified_texfile()
        self.convert_modified_texfile()

if __name__ == "__main__":
    pass
