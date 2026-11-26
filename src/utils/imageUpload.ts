import { translate } from './dictionary.js';

/** The options of an image upload. */
export interface ImageUploadOptions {
    /** The image that it starts with, as an address, or an empty text. */
    value: string;
    /** The accessible name of the upload. */
    label: string;
    /** Called when Done is pressed. */
    ondone: () => void;
}

/** The element of an image upload, which keeps the upload. */
export type ImageUploadElement = HTMLElement & { upload?: ImageUpload };

/**
 * An image upload: the image, and the buttons to choose an image file, to remove the image and to close. An image file
 * can also be dropped on it. The image is kept as a data address, such as "data:image/png;base64,...".
 *
 * It is a dialog: Tab moves between its buttons and stays in it, and Enter and Space press the button that has the focus.
 * The element keeps the upload as its upload property, and the upload has the classes of the jSuites image upload.
 */
export class ImageUpload {
    /** The element of the upload. */
    readonly element: HTMLElement;

    private readonly options: ImageUploadOptions;
    private readonly preview: HTMLElement;
    private readonly input: HTMLInputElement;
    private readonly buttons: HTMLButtonElement[];
    private value: string;

    /**
     * Create an image upload in an element.
     * @param element - The element, which gets the image and the buttons.
     * @param options - The options.
     */
    constructor(element: ImageUploadElement, options: ImageUploadOptions) {
        this.element = element;
        element.upload = this;
        this.options = options;
        this.value = options.value;

        element.classList.add('jupload');
        element.setAttribute('role', 'dialog');
        element.setAttribute('aria-label', options.label);

        this.preview = document.createElement('div');
        this.preview.className = 'jupload-image';
        element.appendChild(this.preview);

        this.input = document.createElement('input');
        this.input.type = 'file';
        this.input.accept = 'image/*';
        this.input.hidden = true;
        this.input.tabIndex = -1;
        this.input.addEventListener('change', () => this.readFile(this.input.files?.[0]));
        element.appendChild(this.input);

        const controls = document.createElement('div');
        controls.className = 'jupload-controls';
        this.buttons = [
            this.createButton(translate('Choose an image'), () => this.input.click()),
            this.createButton(translate('Remove'), () => this.setValue('')),
            this.createButton(translate('Done'), () => this.options.ondone()),
        ];
        controls.append(...this.buttons);
        element.appendChild(controls);

        element.addEventListener('keydown', this.keydown);
        element.addEventListener('dragover', (e) => e.preventDefault());
        element.addEventListener('drop', (e) => {
            e.preventDefault();
            this.readFile(e.dataTransfer?.files[0]);
        });

        this.render();
    }

    /** The image, as an address, or an empty text when there is none. */
    getValue() {
        return this.value;
    }

    /**
     * Show an image.
     * @param value - The address of the image, or an empty text for none.
     */
    setValue(value: string) {
        this.value = value;
        this.render();
    }

    /** Move the focus to the first button. */
    focus() {
        this.buttons[0].focus();
    }

    /**
     * Read an image file, and show it.
     * @param file - The file. Another file than an image is ignored.
     */
    private readFile(file: File | undefined) {
        if (!file?.type.startsWith('image/')) {
            return;
        }
        const reader = new FileReader();
        reader.addEventListener('load', () => this.setValue(reader.result as string));
        reader.readAsDataURL(file);
    }

    private render() {
        this.preview.textContent = '';
        if (this.value) {
            const img = document.createElement('img');
            img.src = this.value;
            img.alt = '';
            this.preview.appendChild(img);
        }
        // Remove is disabled without an image
        this.buttons[1].disabled = !this.value;
    }

    private createButton(text: string, onclick: () => void) {
        const button = document.createElement('button');
        button.type = 'button';
        button.textContent = text;
        button.addEventListener('click', onclick);
        return button;
    }

    private keydown = (e: KeyboardEvent) => {
        const target = e.target as HTMLElement;
        if (e.key === 'Tab') {
            // The focus stays in the dialog
            const buttons = this.buttons.filter((button) => !button.disabled);
            const index = buttons.indexOf(target as HTMLButtonElement);
            const next = (index + (e.shiftKey ? -1 : 1) + buttons.length) % buttons.length;
            buttons[index < 0 ? 0 : next].focus();
        } else if ((e.key === 'Enter' || e.key === ' ') && target.tagName === 'BUTTON') {
            (target as HTMLButtonElement).click();
        } else {
            // Escape closes the editor without saving
            return;
        }
        e.preventDefault();
        e.stopPropagation();
    };
}
