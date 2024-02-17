/* eslint-disable @typescript-eslint/no-unused-vars */
/* global console setInterval, clearInterval */

// Define fetchOpenAI in a scope accessible by GPT function
async function fetchOpenAI(input) {
  var apiKey = localStorage.getItem('OpenAI_ApiKey');
  var selectedModel = localStorage.getItem('SelectedModel');
  if (!apiKey || !selectedModel) {
    console.error('API key or model not set.');
    return;
  }

  console.log(`Fetching OpenAI with model: ${selectedModel} and input: ${input}`);
  const response = await fetch(`https://api.openai.com/v1/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      "model": `${selectedModel}`,
      "messages": [
        {
          "role": "system",
          "content": "You are an expert in categorization!"
        },
        {
          "role": "user",
          "content": `${input}`
        }
      ],
      //max_tokens: 150, // You can adjust based on your needs
      temperature: 0.7, // Optional: control the randomness of the output
      top_p: 1, // Optional: sampling parameter
      frequency_penalty: 0, // Optional: decrease the likelihood of repetition
      presence_penalty: 0, // Optional: increase the likelihood of talking about new concepts
      // ... other parameters you might have
    })
  });

  const data = await response.json();
  console.log('OpenAI response:', data);
  if (data.choices && data.choices.length > 0) {
    return data.choices[0].message.content.trim();
  } else {
    console.error('Unexpected response structure:', data);
    return JSON.stringify(data);
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Office is now ready!");
    // Other initialization code...
  } else {
    console.error("Office is not running in Excel.");
  }
});

/**
 * A custom function that communicates with OpenAI's ChatGPT.
 * @customfunction GPT AskGPT a question.
 * @param {string} input The question or task for GPT-3.
 * @returns {string} The answer from GPT-3.
 */
async function GPT(input) {
  return new Promise((resolve, reject) => {
    try {
      fetchOpenAI(input).then(response => {
        resolve(response);
      }).catch(error => {
        reject(`Error fetching from OpenAI: ${error}`);
      });
    } catch (error) {
      console.error(error);
      reject(`Error: ${error}`);
    }
  });
}

// Associate the custom function with the name "GPT"
CustomFunctions.associate("GPT", GPT);
