<script lang="ts">
  import { GoogleGenerativeAI } from '@google/generative-ai';
  import mammoth from 'mammoth';
  import {
    FileUp,
    AlertCircle,
    Check,
    Loader2,
    Sparkles,
    Send,
    Copy,
    Table,
    CheckCircle2,
    XCircle,
    Clock,
    Image,
    FileText,
    Download,
    Save,
    X,
    Bell
  } from '@lucide/svelte';
  import { Button } from '$lib/components/ui/button';
  import { PUBLIC_AI_API_KEY } from '$env/static/public';
  import { Input } from '$lib/components/ui/input';
  import * as XLSX from 'xlsx';
  import { onMount } from 'svelte';
  import * as Dialog from '$lib/components/ui/dialog';

  import { save } from '@tauri-apps/plugin-dialog';
  import { writeFile } from '@tauri-apps/plugin-fs';

  let apiKey = '';
  let userPrompt = `Based on the following text extracted from a document, identify multiple-choice questions with single correct answers. The questions may be in Vietnamese.

For EACH question, please format the output in this specific machine-readable format:

QUESTION_ID: [Sequential number starting from 1 for each identified question]
QUESTION: [Full question text in Vietnamese or English]
OPTIONS:
A. [Option A text]
B. [Option B text]
C. [Option C text]
D. [Option D text]
CORRECT: [Letter of the correct option, e.g., B]
STATUS: [SUCCESS if you are confident about the question format and correct answer, FAILED if anything is unclear]
REASON: [Only provide if STATUS is FAILED. Brief explanation of why parsing failed]
--- (Use three hyphens as a separator between questions)

Important guidelines:
1. Assign a unique sequential number (1, 2, 3, ...) to each identified question as the QUESTION_ID.
2. Extract the full question text exactly as it appears.
3. If the question block has options A, B, C, D (or a, b, c, d) without accompanying True/False indicators, then proceed to extract:
  - The full question text exactly as it appears.
  - The options labeled A, B, C, D (or a, b, c, d) exactly as they appear.
4. Preserve all Vietnamese text exactly as it appears in the original.
5. Process all questions found in the input text that match the A, B, C, D format.
6. Use three hyphens (---) as a separator between each extracted question block.
7. Crucially, check if the options (A, B, C, D or a, b, c, d) are accompanied by True/False indicators (e.g., Đ, S, Đúng, Sai, T, F, True, False). These indicators are often placed immediately after or near the option text.
8. If such True/False indicators are present next to the options, skip this entire question block. Do not extract it or assign it a QUESTION_ID.

Here is the text:
`;

  interface QuestionData {
    id: string;
    question: string;
    options: { [key: string]: string };
    correctAnswer: string;
    status: 'SUCCESS' | 'FAILED';
    reason?: string;
    isMultiAnswer?: boolean;
    selected?: boolean;
    timeInSeconds?: number;
    imageLink?: string;
    explanation?: string;
  }

  let fileInput: HTMLInputElement;
  let fileContent = '';
  let apiResponse = '';
  let parsedQuestions: QuestionData[] = [];
  let successfulQuestions: QuestionData[] = [];
  let failedQuestions: QuestionData[] = [];
  let isLoading = false;
  let error = '';
  let defaultTimeInSeconds = 60;
  let selectedQuestions: QuestionData[] = [];
  let notification = '';
  let showNotification = false;

  let maxChunkSize = 15000;
  let currentChunk = 1;
  let totalChunks = 1;
  let processingChunks = false;
  let allChunksResponse = "";
  let processedText = "";
  let remainingText = "";

  onMount(async () => {
    try {
      const response = await fetch('https://cdn.mncuchiinhuttt.dev/quizizz-ai-gay.txt');
      const text = await response.text();

      if (text && text.trim()) {
        notification = text.trim();
        showNotification = true;
      }
    } catch (error) {
      console.error('Failed to fetch notification:', error);
    }
  });

  async function handleFileSelect(event: Event) {
    const target = event.target as HTMLInputElement;
    const file = target.files?.[0];
    if (!file) {
      fileContent = '';
      error = 'No file selected.';
      return;
    }

    if (!file.name.endsWith('.docx')) {
      error = 'Please select a .docx file. (.doc is not directly supported)';
      target.value = '';
      return;
    }

    error = '';
    isLoading = true;
    fileContent = '';
    parsedQuestions = [];
    successfulQuestions = [];
    failedQuestions = [];
    apiResponse = '';

    try {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer: arrayBuffer });
      fileContent = result.value;
      console.log('DOCX Extracted Text:', fileContent.substring(0, 500) + '...');
    } catch (err) {
      console.error('Error reading docx file:', err);
      error = 'Error processing the .docx file. Is it corrupted or password-protected?';
      fileContent = '';
    } finally {
      isLoading = false;
    }
  }

  async function getAIResponse() {
    if (!apiKey) {
      try {
        if (typeof PUBLIC_AI_API_KEY !== 'undefined' && PUBLIC_AI_API_KEY) {
          apiKey = PUBLIC_AI_API_KEY;
          console.log('Using environment API key');
        } else {
          error = 'Please enter your Google API Key or set it in your environment variables.';
          return;
        }
      } catch (err) {
        error = 'Please enter your Google API Key.';
        return;
      }
    }

    if (!fileContent) {
      error = 'Please select and process a .docx file first.';
      return;
    }

    isLoading = true;
    error = '';
    apiResponse = '';
    parsedQuestions = [];
    successfulQuestions = [];
    failedQuestions = [];
    allChunksResponse = "";

    if (fileContent.length > maxChunkSize) {
      processingChunks = true;
      totalChunks = Math.ceil(fileContent.length / maxChunkSize);
      currentChunk = 1;
      remainingText = fileContent;
      processedText = "";

      while (remainingText.length > 0 && currentChunk <= totalChunks) {
        await processChunk();
      }

      apiResponse = allChunksResponse;
      parseAIResponse(apiResponse);
      processingChunks = false;
      isLoading = false;
    } else {
      await processSingleRequest(fileContent);
    }
  }

  async function processChunk() {
    try {
      const chunkSize = Math.min(maxChunkSize, remainingText.length);

      let breakpoint = chunkSize;
      if (chunkSize < remainingText.length) {
        const breakChar = remainingText.lastIndexOf("\n\n", chunkSize);
        if (breakChar > chunkSize * 0.8) {
          breakpoint = breakChar;
        }
      }

      const currentText = remainingText.substring(0, breakpoint);

      let chunkPrompt = userPrompt;
      if (currentChunk > 1) {
        chunkPrompt = `This is CHUNK ${currentChunk} of ${totalChunks} from the document. Continue parsing multiple-choice questions from where you left off.
        
For EACH question, please format the output in this specific machine-readable format:

QUESTION_ID: [Sequential number continuing from previous chunks]
QUESTION: [Full question text in Vietnamese or English]
OPTIONS:
A. [Option A text]
B. [Option B text]
C. [Option C text]
D. [Option D text]
CORRECT: [Letter of the correct option, e.g., B]
STATUS: [SUCCESS if you are confident about the question format and correct answer, FAILED if anything is unclear]
REASON: [Only provide if STATUS is FAILED. Brief explanation of why parsing failed]
--- (Use three hyphens as a separator between questions)

Important guidelines:
1. Assign a unique sequential number (1, 2, 3, ...) to each identified question as the QUESTION_ID.
2. Extract the full question text exactly as it appears.
3. Extract the options labeled A, B, C, D exactly as they appear.
4. Preserve all Vietnamese text exactly as it appears in the original.
5. Process all questions found in the input text that match the A, B, C, D format.
6. Use three hyphens (---) as a separator between each extracted question block.

Here is the text for this chunk:
        `;
      }

      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

      const fullPrompt = chunkPrompt + '\n\n' + currentText;

      const result = await model.generateContent(fullPrompt);
      const response = result.response;
      const chunkResponse = response.text();

      allChunksResponse += (currentChunk > 1 ? "\n" : "") + chunkResponse;

      processedText += currentText;
      remainingText = remainingText.substring(breakpoint);

      currentChunk++;
    } catch (err: any) {
      console.error(`Error processing chunk ${currentChunk}:`, err);
      error = `Error processing chunk ${currentChunk}: ${err.message || 'Unknown error'}`;
      remainingText = "";
    }
  }

  async function processSingleRequest(content: string) {
    try {
      const genAI = new GoogleGenerativeAI(apiKey);
      const model = genAI.getGenerativeModel({ model: 'gemini-1.5-flash' });

      const fullPrompt = userPrompt + '\n\n' + content;

      const result = await model.generateContent(fullPrompt);
      const response = result.response;
      apiResponse = response.text();

      console.log('AI Raw Response:', apiResponse);
      parseAIResponse(apiResponse);
    } catch (err: any) {
      console.error('Google AI API Error:', err);
      if (err.message?.includes('API key not valid')) {
        error =
          'API Key not valid. Please check your key and ensure the Gemini API is enabled in your Google Cloud project.';
      } else if (err.message?.includes('429')) {
        error = 'API Rate Limit Exceeded. Please wait and try again later.';
      } else if (err.message?.includes('fetch')) {
        error = 'Network error connecting to Google AI. Check your internet connection.';
      } else {
        error = 'Error fetching response from Google AI: ' + (err.message || 'Unknown error');
      }
      apiResponse = '';
    } finally {
      isLoading = false;
    }
  }

  function parseAIResponse(text: string) {
    const regex =
      /QUESTION_ID:\s*(\d+|[A-Za-z0-9-]+)\s*\n?QUESTION:\s*([\s\S]*?)(?=\nOPTIONS:)\nOPTIONS:\s*\n?((?:(?:[A-Z]\.\s*[\s\S]*?)(?=\n[A-Z]\.|CORRECT:|$))*)\s*CORRECT:\s*([A-Z]|Unknown)\s*\n?STATUS:\s*(SUCCESS|FAILED)\s*\n?(?:REASON:\s*([\s\S]*?))?(?=\n---|\n\n|$)/g;

    const questions: QuestionData[] = [];
    const successful: QuestionData[] = [];
    const failed: QuestionData[] = [];

    try {
      let currentMatch;
      const matches = [];

      while ((currentMatch = regex.exec(text)) !== null) {
        matches.push(currentMatch);
      }

      if (matches.length === 0) {
        console.warn(
          'No matches found with regex pattern. First 300 chars of raw text:',
          text.substring(0, 300) + '...'
        );
        error =
          'Could not parse any questions from the AI response. The format might be different than expected.';
      }

      console.log(`Found ${matches.length} question matches in the response`);

      for (const match of matches) {
        const [_, id, questionText, optionsRaw, correctAnswer, status, reason] = match;

        const optionsRegex = /([A-Z])\.\s*([\s\S]*?)(?=\n[A-Z]\.|$)/g;
        const options: { [key: string]: string } = {};

        let optionMatch;
        while ((optionMatch = optionsRegex.exec(optionsRaw)) !== null) {
          const [__, key, optionText] = optionMatch;
          options[key] = optionText.trim();
        }

        const questionData: QuestionData = {
          id: id.trim(),
          question: questionText.trim(),
          options: options,
          correctAnswer: correctAnswer.trim(),
          status: (status?.trim().toUpperCase() || 'FAILED') as 'SUCCESS' | 'FAILED',
          reason: reason ? reason.trim() : undefined,
          isMultiAnswer: false
        };

        if (Object.keys(questionData.options).length > 0) {
          questions.push(questionData);

          if (questionData.status === 'SUCCESS') {
            successful.push(questionData);
          } else {
            failed.push(questionData);
          }
        } else {
          console.warn(`Question ${id} skipped - no options found`);
        }
      }

      const vnRegex =
        /Câu\s*(\d+)[.\s:]+([\s\S]*?)(?=A\.|\([Aa]\))([Aa][\.\)][\s\S]*?)(?=B\.|\([Bb]\))([Bb][\.\)][\s\S]*?)(?=C\.|\([Cc]\))([Cc][\.\)][\s\S]*?)(?=D\.|\([Dd]\))([Dd][\.\)][\s\S]*?)(?=Câu\s*\d+|$)/g;

      const vnMatches = [];
      let vnMatch: RegExpExecArray | null;
      while ((vnMatch = vnRegex.exec(text)) !== null) {
        const questionId = vnMatch[1]?.trim();
        if (questionId) {
          const existingQuestion = questions.find((q) => q.id === questionId);
          if (!existingQuestion) {
            vnMatches.push(vnMatch);
          }
        } else {
          console.warn("VN Regex matched but couldn't capture ID:", vnMatch[0]);
        }
      }

      if (vnMatches.length > 0) {
        console.log(`Found ${vnMatches.length} potential Vietnamese questions missed by primary regex.`);
      }

      for (let i = 0; i < vnMatches.length; i++) {
        const [_, id, questionText, optA, optB, optC, optD]: (string | undefined)[] = vnMatches[i];

        if (id && questionText && optA && optB && optC && optD) {
          const q: QuestionData = {
            id: id.trim(),
            question: questionText.trim(),
            options: {
              A: optA.replace(/^[Aa][\.\)]\s*/, '').trim(),
              B: optB.replace(/^[Bb][\.\)]\s*/, '').trim(),
              C: optC.replace(/^[Cc][\.\)]\s*/, '').trim(),
              D: optD.replace(/^[Dd][\.\)]\s*/, '').trim()
            },
            correctAnswer: 'Unknown',
            status: 'FAILED',
            reason: 'Could not reliably determine the correct answer from this format.',
            isMultiAnswer: false
          };

          questions.push(q);
          failed.push(q);
        } else {
          console.warn('Skipping Vietnamese match due to missing parts:', vnMatches[i]);
        }
      }

      if (questions.length === 0 && text.trim().length > 0) {
        console.warn('No valid questions were created. Check the AI output format.');
        error = 'Could not parse valid questions from the AI response. Check the raw response below.';
      }
    } catch (err) {
      console.error('Error parsing AI response:', err);
      error = 'Error parsing the AI response: ' + (err instanceof Error ? err.message : String(err));
    }

    parsedQuestions = questions;
    successfulQuestions = successful;
    failedQuestions = failed;

    console.log('Parsed Questions:', parsedQuestions);
    console.log('Successful Questions:', successfulQuestions);
    console.log('Failed Questions:', failedQuestions);
  }

  function fadeIn(node: HTMLElement, { delay = 0, duration = 300 }) {
    return {
      delay,
      duration,
      css: (t: number) => `opacity: ${t}; transform: translateY(${(1 - t) * 10}px)`
    };
  }

  function updateSelectedQuestions() {
    selectedQuestions = successfulQuestions.filter((q: QuestionData) => q.selected);
  }

  function applyDefaultTime() {
    successfulQuestions.forEach((q: QuestionData) => {
      if (q.selected) {
        q.timeInSeconds = defaultTimeInSeconds;
      }
    });
    successfulQuestions = [...successfulQuestions];
    updateSelectedQuestions();
  }

  async function exportToExcel() {
    if (selectedQuestions.length === 0) {
      error = 'Please select at least one question to export.';
      return;
    }

    const workbookData = selectedQuestions.map((q: QuestionData) => {
      const optionToNumber: { [key: string]: number } = {
        A: 1,
        B: 2,
        C: 3,
        D: 4,
        E: 5
      };

      const optionKeys = Object.keys(q.options);
      const options = optionKeys.map((key) => q.options[key]);

      let correctAnswer = '';
      if (q.isMultiAnswer) {
        correctAnswer = q.correctAnswer
          .split(',')
          .map((ans) => optionToNumber[ans])
          .join(', ');
      } else {
        correctAnswer = String(optionToNumber[q.correctAnswer] || '');
      }

      return {
        'Question Text': q.question,
        'Question Type': q.isMultiAnswer ? 'Checkbox' : 'Multiple Choice',
        'Option 1': options[0] || '',
        'Option 2': options[1] || '',
        'Option 3': options[2] || '',
        'Option 4': options[3] || '',
        'Option 5': options[4] || '',
        'Correct Answer': correctAnswer,
        'Time in seconds': q.timeInSeconds || defaultTimeInSeconds,
        'Image Link': q.imageLink || '',
        'Answer explanation': q.explanation || ''
      };
    });

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(workbookData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Questions');

    try {
      const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

      const filePath = await save({
        defaultPath: `quizizz-questions-${(new Date()).toISOString().split('T')[0]}.xlsx`,
        filters: [{
          name: 'Excel',
          extensions: ['xlsx']
        }]
      });

      if (filePath) {
        await writeFile(filePath, excelBuffer);
        const successDialog = document.getElementById('success-dialog');
        if (successDialog) successDialog.classList.remove('hidden');
      }
    } catch (err) {
      console.error('Error saving file:', err);
      error = `Error saving file: ${err instanceof Error ? err.message : String(err)}`;
    }
  }

  function closeNotification() {
    showNotification = false;
  }
</script>

<div class="min-h-screen bg-gradient-to-br from-[#f8fafc] to-[#e9eef8] dark:from-slate-900 dark:to-slate-800 p-4 sm:p-6">
  <Dialog.Root bind:open={showNotification}>
    <Dialog.Portal>
      <Dialog.Overlay />
      <Dialog.Content class="max-w-lg">
        <Dialog.Header>
          <Dialog.Title class="flex items-center gap-2">
            <Bell class="h-5 w-5 text-indigo-500" />
            Notification
          </Dialog.Title>
        </Dialog.Header>
        <div class="whitespace-pre-wrap text-slate-700 dark:text-slate-300 max-h-[60vh] overflow-y-auto scrollbar">
          {notification}
        </div>
        <Dialog.Footer>
          <Dialog.Close asChild>
            <Button onclick={closeNotification}>Close</Button>
          </Dialog.Close>
        </Dialog.Footer>
      </Dialog.Content>
    </Dialog.Portal>
  </Dialog.Root>

  <main class="max-w-5xl mx-auto bg-white dark:bg-slate-900 rounded-xl shadow-xl overflow-hidden">
    <div class="p-6 sm:p-8 border-b border-slate-200 dark:border-slate-700 bg-gradient-to-r from-indigo-500 to-purple-600 text-white">
      <h1 class="text-3xl sm:text-4xl font-bold text-center mb-2">
        DOCX Question Extractor <span class="text-yellow-300">AI</span>
      </h1>
      <p class="text-center text-indigo-100 max-w-2xl mx-auto">Extract multiple-choice questions from DOCX files using AI to create quizzes effortlessly</p>
    </div>

    <section class="p-6 sm:p-8">
      <div class="space-y-6">
        <div class="rounded-lg border border-slate-200 dark:border-slate-700 overflow-hidden">
          <div class="bg-slate-50 dark:bg-slate-800 px-4 py-3 border-b border-slate-200 dark:border-slate-700">
            <h2 class="font-medium text-slate-700 dark:text-slate-300 flex items-center gap-2">
              <Sparkles class="h-4 w-4 text-indigo-500" />
              Configuration
            </h2>
          </div>
          
          <div class="p-5 space-y-5">
            <div class="flex flex-col sm:flex-row gap-4">
              <div class="sm:w-1/2">
                <label for="apiKey" class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1.5">Google API Key</label>
                <div class="relative">
                  <input 
                    type="password" 
                    id="apiKey" 
                    bind:value={apiKey} 
                    placeholder="Enter your Google AI API Key or leave blank to use environment key" 
                    class="w-full px-3 py-2.5 border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 dark:focus:ring-indigo-400 focus:border-indigo-500 dark:focus:border-indigo-400 transition-all text-slate-800 dark:text-slate-200 placeholder-slate-400 dark:placeholder-slate-500" 
                  />
                </div>
                <p class="mt-1.5 text-xs text-slate-500 dark:text-slate-400">Leave blank to use environment API key if available, or enter your own key.</p>
              </div>
              
              <div class="sm:w-1/2">
                <label for="chunkSize" class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1.5">Chunk Size</label>
                <div class="relative">
                  <input 
                    type="number" 
                    id="chunkSize" 
                    bind:value={maxChunkSize} 
                    min="5000" 
                    max="25000"
                    step="5000"
                    class="w-full px-3 py-2.5 border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 dark:focus:ring-indigo-400 focus:border-indigo-500 dark:focus:border-indigo-400 transition-all text-slate-800 dark:text-slate-200 placeholder-slate-400 dark:placeholder-slate-500" 
                  />
                </div>
                <p class="mt-1.5 text-xs text-slate-500 dark:text-slate-400">Characters per chunk (5,000-25,000). Lower is safer but slower.</p>
              </div>
            </div>
            
            <div>
              <label for="fileInput" class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1.5">Upload DOCX File</label>
              <div class="mt-1 flex justify-center px-6 pt-5 pb-6 border-2 border-dashed border-slate-300 dark:border-slate-600 rounded-lg hover:border-indigo-400 dark:hover:border-indigo-500 transition-colors cursor-pointer" 
                   on:click={() => fileInput.click()}
              >
                <div class="space-y-2 text-center">
                  <FileUp class="mx-auto h-12 w-12 text-slate-400" />
                  <div class="flex text-sm text-slate-600 dark:text-slate-400">
                    <label for="fileInput" class="relative cursor-pointer rounded-md font-medium text-indigo-600 dark:text-indigo-400 hover:text-indigo-500 focus-within:outline-none">
                      <span>Upload a file</span>
                      <input id="fileInput" bind:this={fileInput} on:change={handleFileSelect} type="file" accept=".docx" class="sr-only" />
                    </label>
                    <p class="pl-1">or drag and drop</p>
                  </div>
                  <p class="text-xs text-slate-500 dark:text-slate-400">
                    Only .docx files are supported
                  </p>
                </div>
              </div>
            </div>
            
            {#if fileContent && !error && !isLoading}
              <div class="flex items-center p-3 bg-green-50 dark:bg-green-900/20 border border-green-200 dark:border-green-800 rounded-md text-sm text-green-700 dark:text-green-400 animate-in fade-in duration-300 ease-out slide-in-from-top-2">
                <Check class="h-5 w-5 mr-2 flex-shrink-0" />
                <span>File loaded successfully. Found approx. {fileContent.length} characters. Ready to extract.</span>
              </div>
            {/if}

            <Button
              onclick={getAIResponse}
              disabled={isLoading || !fileContent}
              class="w-full flex justify-center items-center px-6 py-3 border border-transparent text-base font-medium rounded-md shadow-sm text-white bg-indigo-600 hover:bg-indigo-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 disabled:bg-slate-400 disabled:cursor-not-allowed transition-all duration-200"
            >
              {#if isLoading}
                <Loader2 class="animate-spin -ml-1 mr-3 h-5 w-5 text-white" />
                Processing...
              {:else}
                <Send class="mr-2 h-5 w-5" />
                Extract Questions
              {/if}
            </Button>
          </div>
        </div>
      </div>
    </section>

    {#if error}
      <div class="mx-6 sm:mx-8 mb-6 p-4 bg-red-50 dark:bg-red-900/20 border border-red-200 dark:border-red-800 rounded-md text-sm text-red-700 dark:text-red-400 flex items-start" role="alert">
        <AlertCircle class="h-5 w-5 mr-2 flex-shrink-0" />
        <span><strong class="font-semibold">Error:</strong> {error}</span>
      </div>
    {/if}

    {#if isLoading}
      <div class="my-10 text-center text-slate-600 dark:text-slate-400">
        <div class="inline-block h-16 w-16 animate-spin rounded-full border-4 border-solid border-indigo-600 dark:border-indigo-400 border-r-transparent align-[-0.125em] motion-reduce:animate-[spin_1.5s_linear_infinite]" role="status">
          <span class="!absolute !-m-px !h-px !w-px !overflow-hidden !whitespace-nowrap !border-0 !p-0 ![clip:rect(0,0,0,0)]">Loading...</span>
        </div>
        <p class="mt-4 text-lg font-medium">Analyzing document with AI...</p>
        {#if processingChunks}
          <div class="mt-2 flex flex-col items-center">
            <p class="text-sm text-slate-500 dark:text-slate-400">
              Processing chunk {currentChunk - 1} of {totalChunks}
            </p>
            <div class="w-64 h-2 mt-2 bg-slate-200 dark:bg-slate-700 rounded-full overflow-hidden">
              <div class="h-full bg-indigo-500 transition-all duration-300" style="width: {Math.round(((currentChunk - 1) / totalChunks) * 100)}%"></div>
            </div>
          </div>
        {:else}
          <p class="text-sm text-slate-500 dark:text-slate-400">This may take a moment depending on document size</p>
        {/if}
      </div>
    {/if}

    {#if successfulQuestions.length > 0 || failedQuestions.length > 0 || apiResponse}
    <section class="px-6 sm:px-8 pb-8 pt-2">
      <div class="mb-8 flex flex-wrap gap-4">
        <div class="flex-1 min-w-[180px] bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4 shadow-sm">
          <h3 class="text-sm font-medium text-slate-500 dark:text-slate-400 mb-1">Total Questions</h3>
          <div class="flex items-center">
            <span class="text-3xl font-bold text-slate-800 dark:text-slate-200">{parsedQuestions.length}</span>
            <div class="ml-auto rounded-full bg-slate-100 dark:bg-slate-700 p-2">
              <Table class="h-6 w-6 text-indigo-500" />
            </div>
          </div>
        </div>
        
        <div class="flex-1 min-w-[180px] bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4 shadow-sm">
          <h3 class="text-sm font-medium text-slate-500 dark:text-slate-400 mb-1">Successfully Extracted</h3>
          <div class="flex items-center">
            <span class="text-3xl font-bold text-green-600 dark:text-green-400">{successfulQuestions.length}</span>
            <div class="ml-auto rounded-full bg-green-100 dark:bg-green-900/30 p-2">
              <CheckCircle2 class="h-6 w-6 text-green-500" />
            </div>
          </div>
        </div>
        
        <div class="flex-1 min-w-[180px] bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-4 shadow-sm">
          <h3 class="text-sm font-medium text-slate-500 dark:text-slate-400 mb-1">Failed to Extract</h3>
          <div class="flex items-center">
            <span class="text-3xl font-bold text-red-600 dark:text-red-400">{failedQuestions.length}</span>
            <div class="ml-auto rounded-full bg-red-100 dark:bg-red-900/30 p-2">
              <XCircle class="h-6 w-6 text-red-500" />
            </div>
          </div>
        </div>
      </div>

      {#if successfulQuestions.length > 0}
        <div class="mb-8 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700 p-5">
          <div class="flex items-center justify-between mb-4">
            <h2 class="text-xl font-semibold text-slate-800 dark:text-slate-200 flex items-center">
              <Download class="h-5 w-5 mr-2 text-indigo-500" />
              Export Questions
            </h2>
            <div class="flex gap-2">
              <span class="text-sm text-slate-600 dark:text-slate-400">
                {selectedQuestions.length} of {successfulQuestions.length} selected
              </span>
              <Button 
                variant="outline" 
                size="sm" 
                onclick={() => {
                  successfulQuestions.forEach(q => q.selected = true);
                  updateSelectedQuestions();
                }}
              >
                Select all
              </Button>
              <Button 
                variant="outline" 
                size="sm" 
                onclick={() => {
                  successfulQuestions.forEach(q => q.selected = false);
                  updateSelectedQuestions();
                }}
              >
                Clear all
              </Button>
            </div>
          </div>
          
          <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
            <div>
              <label class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1.5">
                Default Time (seconds)
              </label>
              <div class="flex items-center">
                <Input 
                  type="number" 
                  bind:value={defaultTimeInSeconds} 
                  min="1" 
                  class="w-32 mr-2"
                />
                <Button
                  variant="secondary"
                  size="sm"
                  onclick={applyDefaultTime}
                  disabled={selectedQuestions.length === 0}
                >
                  <Clock class="h-4 w-4 mr-1.5" />
                  Apply to selected
                </Button>
              </div>
              <p class="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Time limit for each question
              </p>
            </div>
            
            <div class="col-span-2">
              <Button
                onclick={exportToExcel}
                disabled={selectedQuestions.length === 0}
                class="bg-green-600 hover:bg-green-700 text-white"
              >
                <Save class="h-4 w-4 mr-2" />
                Save {selectedQuestions.length} questions to Excel
              </Button>
              <p class="mt-1 text-xs text-slate-500 dark:text-slate-400">
                Creates XLSX file with selected questions formatted for Quizizz import
              </p>
            </div>
          </div>
        </div>
      {/if}

      {#if successfulQuestions.length > 0}
        <div class="mb-8">
          <div class="flex items-center justify-between mb-4">
            <h2 class="text-xl font-semibold text-slate-800 dark:text-slate-200 flex items-center">
              <CheckCircle2 class="h-5 w-5 mr-2 text-green-500" />
              Successfully Extracted Questions
            </h2>
          </div>
          
          <div class="overflow-x-auto">
            <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700">
              <thead>
                <tr>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider w-10">
                    <span class="sr-only">Select</span>
                  </th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">ID</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Question</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Options</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Correct</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Type</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Actions</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200 dark:divide-slate-700">
                {#each successfulQuestions as q, i (i)}
                  <tr transition:fadeIn={{delay: i * 30}} class="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td class="px-4 py-3 whitespace-nowrap">
                      <input 
                        type="checkbox" 
                        bind:checked={q.selected} 
                        on:change={() => updateSelectedQuestions()}
                        class="h-4 w-4 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500 dark:border-slate-600 dark:bg-slate-700 dark:checked:bg-indigo-600"
                      />
                    </td>
                    <td class="px-4 py-3 whitespace-nowrap text-sm font-medium text-slate-800 dark:text-slate-200">{q.id}</td>
                    <td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300 max-w-[300px] truncate">{q.question}</td>
                    <td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300">
                      <div class="flex flex-col gap-1">
                        {#each Object.entries(q.options) as [key, value]}
                          <div class="flex items-start">
                            <span class={`font-medium mr-1 ${q.correctAnswer.split(',').includes(key) ? 'text-green-600 dark:text-green-400' : ''}`}>{key}:</span> 
                            <span class="truncate max-w-[200px]">{value}</span>
                          </div>
                        {/each}
                      </div>
                    </td>
                    <td class="px-4 py-3 whitespace-nowrap">
                      {#if q.correctAnswer === 'Unknown'}
                        <span class="px-2 py-1 text-xs font-medium rounded-full bg-yellow-100 dark:bg-yellow-900/30 text-yellow-800 dark:text-yellow-400">
                          Unknown
                        </span>
                      {:else if q.isMultiAnswer}
                        <div class="flex flex-wrap gap-1">
                          {#each q.correctAnswer.split(',') as answer}
                            <span class="px-2 py-1 text-xs font-medium rounded-full bg-green-100 dark:bg-green-900/30 text-green-800 dark:text-green-400">
                              {answer}
                            </span>
                          {/each}
                        </div>
                      {:else}
                        <span class="px-2 py-1 text-xs font-medium rounded-full bg-green-100 dark:bg-green-900/30 text-green-800 dark:text-green-400">
                          {q.correctAnswer}
                        </span>
                      {/if}
                    </td>
                    <td class="px-4 py-3 whitespace-nowrap text-xs">
                      {#if q.isMultiAnswer}
                        <span class="px-2 py-1 text-xs font-medium rounded-full bg-blue-100 dark:bg-blue-900/30 text-blue-800 dark:text-blue-400">
                          Multi
                        </span>
                      {:else}
                        <span class="px-2 py-1 text-xs font-medium rounded-full bg-purple-100 dark:bg-purple-900/30 text-purple-800 dark:text-purple-400">
                          Single
                        </span>
                      {/if}
                    </td>
                    <td class="px-4 py-3 whitespace-nowrap text-sm">
                      <button 
                        class="p-1.5 rounded-md text-slate-500 hover:text-slate-700 hover:bg-slate-100 dark:text-slate-400 dark:hover:text-slate-200 dark:hover:bg-slate-700 transition-colors"
                        title="Edit Details"
                        on:click={() => {
                          const modal = document.getElementById(`edit-modal-${i}`);
                          if (modal) modal.classList.remove('hidden');
                        }}
                      >
                        <FileText class="h-4 w-4" />
                      </button>
                      
                      <div id={`edit-modal-${i}`} class="fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4 hidden">
                        <div class="bg-white dark:bg-slate-800 rounded-lg shadow-xl max-w-lg w-full max-h-[90vh] overflow-auto">
                          <div class="p-4 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
                            <h3 class="text-lg font-medium text-slate-800 dark:text-slate-200">Edit Question #{q.id}</h3>
                            <button
                              class="text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200"
                              on:click={() => {
                                const modal = document.getElementById(`edit-modal-${i}`);
                                if (modal) modal.classList.add('hidden');
                              }}
                            >
                              &times;
                            </button>
                          </div>
                          <div class="p-4 space-y-4">
                            <div>
                              <label class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                                Time in seconds
                              </label>
                              <Input 
                                type="number" 
                                bind:value={q.timeInSeconds} 
                                min="1" 
                                class="w-full"
                              />
                            </div>
                            <div>
                              <label class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                                Image Link (optional)
                              </label>
                              <div class="flex gap-2 items-center">
                                <Input 
                                  type="text" 
                                  bind:value={q.imageLink} 
                                  placeholder="https://example.com/image.jpg" 
                                  class="w-full"
                                />
                                <Image class="h-4 w-4 text-slate-400" />
                              </div>
                            </div>
                            <div>
                              <label class="block text-sm font-medium text-slate-700 dark:text-slate-300 mb-1">
                                Answer Explanation (optional)
                              </label>
                              <textarea
                                bind:value={q.explanation}
                                placeholder="Explain why this answer is correct..."
                                class="w-full h-24 px-3 py-2 border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 rounded-md shadow-sm focus:outline-none focus:ring-2 focus:ring-indigo-500 dark:focus:ring-indigo-400 focus:border-indigo-500 dark:focus:border-indigo-400 text-slate-800 dark:text-slate-200 placeholder-slate-400 dark:placeholder-slate-500"
                              ></textarea>
                            </div>
                          </div>
                          <div class="p-4 border-t border-slate-200 dark:border-slate-700 flex justify-end gap-2">
                            <Button
                              variant="outline"
                              onclick={() => {
                                const modal = document.getElementById(`edit-modal-${i}`);
                                if (modal) modal.classList.add('hidden');
                              }}
                            >
                              Cancel
                            </Button>
                            <Button
                              onclick={() => {
                                const modal = document.getElementById(`edit-modal-${i}`);
                                if (modal) modal.classList.add('hidden');
                                if (!q.selected) {
                                  q.selected = true;
                                  updateSelectedQuestions();
                                }
                              }}
                            >
                              Save
                            </Button>
                          </div>
                        </div>
                      </div>
                    </td>
                  </tr>
                {/each}
              </tbody>
            </table>
          </div>
        </div>
      {/if}

      {#if failedQuestions.length > 0}
        <div class="mb-8">
          <div class="flex items-center justify-between mb-4">
            <h2 class="text-xl font-semibold text-slate-800 dark:text-slate-200 flex items-center">
              <XCircle class="h-5 w-5 mr-2 text-red-500" />
              Failed Questions
            </h2>
          </div>
          
          <div class="overflow-x-auto">
            <table class="min-w-full divide-y divide-slate-200 dark:divide-slate-700 bg-white dark:bg-slate-800 rounded-lg border border-slate-200 dark:border-slate-700">
              <thead>
                <tr>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">ID</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Question</th>
                  <th scope="col" class="px-4 py-3 text-left text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Failure Reason</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-slate-200 dark:divide-slate-700">
                {#each failedQuestions as q, i (i)}
                  <tr transition:fadeIn={{delay: i * 30}} class="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td class="px-4 py-3 whitespace-nowrap text-sm font-medium text-slate-800 dark:text-slate-200">{q.id}</td>
                    <td class="px-4 py-3 text-sm text-slate-700 dark:text-slate-300 max-w-[300px] truncate">{q.question}</td>
                    <td class="px-4 py-3 text-sm text-red-600 dark:text-red-400">{q.reason || 'Unknown error'}</td>
                  </tr>
                {/each}
              </tbody>
            </table>
          </div>
        </div>
      {/if}

      {#if apiResponse}
        <div class="mt-10 pt-6 border-t border-slate-200 dark:border-slate-700">
            <div class="flex items-center justify-between mb-3">
              <h3 class="text-lg font-medium text-slate-700 dark:text-slate-300">Raw AI Response</h3>
              <button 
                class="inline-flex items-center px-2.5 py-1.5 border border-slate-300 dark:border-slate-600 shadow-sm text-xs font-medium rounded text-slate-700 dark:text-slate-300 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-indigo-500 transition-colors"
                on:click={() => navigator.clipboard.writeText(apiResponse)}
              >
                <Copy class="h-3.5 w-3.5 mr-1.5" />
                Copy
              </button>
            </div>
            <div class="bg-slate-900 text-slate-200 p-4 rounded-lg text-sm overflow-x-auto whitespace-pre-wrap break-words max-h-[400px] overflow-y-auto scrollbar">
              {apiResponse}
            </div>
        </div>
      {/if}

    </section>
    {/if}

    <div id="success-dialog" class="hidden fixed inset-0 bg-black bg-opacity-50 z-50 flex items-center justify-center p-4">
      <div class="bg-white dark:bg-slate-800 rounded-lg shadow-xl max-w-lg w-full max-h-[90vh] overflow-auto">
        <div class="p-4 border-b border-slate-200 dark:border-slate-700 flex justify-between items-center">
          <h3 class="text-lg font-medium text-slate-800 dark:text-slate-200 flex items-center gap-2">
            <Check class="h-5 w-5 text-green-500" />
            Success
          </h3>
          <button
            class="text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200"
            on:click={() => {
              const modal = document.getElementById('success-dialog');
              if (modal) modal.classList.add('hidden');
            }}
          >
            <X class="h-5 w-5" />
          </button>
        </div>
        <div class="p-6 text-slate-700 dark:text-slate-300">
          <p>Excel file successfully saved!</p>
        </div>
        <div class="p-4 border-t border-slate-200 dark:border-slate-700 flex justify-end">
          <Button
            onclick={() => {
              const modal = document.getElementById('success-dialog');
              if (modal) modal.classList.add('hidden');
            }}
          >
            Close
          </Button>
        </div>
      </div>
    </div>
  </main>
  
  <footer class="mt-6 text-center text-sm text-slate-500 dark:text-slate-400">
    <p>© {new Date().getFullYear()} DOCX Question Extractor AI • Made with SvelteKit & Tailwind CSS</p>
    
    <div class="mt-4 pt-4 border-t border-slate-200 dark:border-slate-700 flex flex-col md:flex-row items-center justify-center gap-4">
      <div class="flex items-center gap-3">
        <img 
          src="https://avatars.githubusercontent.com/u/76028730?v=4" 
          alt="Developer Avatar" 
          class="w-10 h-10 rounded-full border-2 border-indigo-500"
        />
        <div class="text-left">
          <h3 class="font-medium text-slate-700 dark:text-slate-300">Vo Minh Long</h3>
          <p class="text-xs text-slate-500 dark:text-slate-400">Full-stack Developer</p>
        </div>
      </div>
      
      <div class="flex items-center gap-4">
        <a 
          href="https://github.com/mncuchiinhuttt/quizizz-ai-import" 
          target="_blank" 
          rel="noopener noreferrer" 
          class="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-md border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-800 hover:bg-slate-50 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 text-sm transition-colors"
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="lucide lucide-github">
            <path d="M15 22v-4a4.8 4.8 0 0 0-1-3.5c3 0 6-2 6-5.5.08-1.25-.27-2.48-1-3.5.28-1.15.28-2.35 0-3.5 0 0-1 0-3 1.5-2.64-.5-5.36-.5-8 0C6 2 5 2 5 2c-.3 1.15-.3 2.35 0 3.5A5.403 5.403 0 0 0 4 9c0 3.5 3 5.5 6 5.5-.39.49-.68 1.05-.85 1.65-.17.6-.22 1.23-.15 1.85v4"></path>
            <path d="M9 18c-4.51 2-5-2-7-2"></path>
          </svg>
          GitHub Repository
        </a>
        
      </div>
    </div>
  </footer>
</div>

<style>
  .scrollbar::-webkit-scrollbar {
    width: 8px;
    height: 8px;
  }
  
  .scrollbar::-webkit-scrollbar-track {
    background: rgba(0, 0, 0, 0.1);
    border-radius: 4px;
  }
  
  .scrollbar::-webkit-scrollbar-thumb {
    background: rgba(255, 255, 255, 0.2);
    border-radius: 4px;
  }
  
  .scrollbar::-webkit-scrollbar-thumb:hover {
    background: rgba(255, 255, 255, 0.3);
  }
</style>
