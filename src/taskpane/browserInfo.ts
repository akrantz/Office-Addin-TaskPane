import * as bowser from "bowser";

/* global navigator */

export class BrowserInfo
{
    private name: string;
    private version: string;

    get Name(): string
    { 
        return this.name;
    }
    
    get Version(): string
    {
        return this.version;
    }

    public constructor(userAgent?: string)
    {
        if (!userAgent) {
            if (navigator != null) {
                userAgent = navigator.userAgent
            }
        }

        if (userAgent) {
            const result = bowser.getParser(userAgent).getResult();

            this.name = result.browser.name;
            this.version = result.browser.version;
        }
    }
}