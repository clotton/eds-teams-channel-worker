export function isQuestion(text) {
    if (!text) return false;
    const lower = text.toLowerCase().trim();
    const questionWords = ['who', 'what', 'when', 'where', 'why', 'how'];

    // Check if text ends with '?'
    if (lower.endsWith('?')) return true;

    // Or contains question words near start (optional)
    return questionWords.some(word => lower.startsWith(word + ' '));
}
