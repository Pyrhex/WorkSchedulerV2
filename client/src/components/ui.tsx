import type {
  ButtonHTMLAttributes,
  InputHTMLAttributes,
  PropsWithChildren,
  Ref,
  SelectHTMLAttributes,
  TextareaHTMLAttributes,
} from 'react';

import { clsx } from 'clsx';

export function Card({
  children,
  className,
  ref,
}: PropsWithChildren<{ className?: string; ref?: Ref<HTMLElement> }>) {
  return (
    <section
      ref={ref}
      className={clsx(
        'rounded-[28px] border border-white/60 bg-white/90 p-5 shadow-panel backdrop-blur md:p-6',
        className,
      )}
    >
      {children}
    </section>
  );
}

export function Button({
  className,
  tone = 'primary',
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & {
  tone?: 'primary' | 'secondary' | 'danger';
}) {
  return (
    <button
      className={clsx(
        'inline-flex items-center justify-center rounded-full px-4 py-2 text-sm font-semibold transition disabled:cursor-not-allowed disabled:opacity-50',
        tone === 'primary' && 'bg-ink text-white hover:bg-pine',
        tone === 'secondary' &&
          'border border-ink/10 bg-mist text-ink hover:border-fairway hover:text-fairway',
        tone === 'danger' && 'bg-clay text-white hover:bg-[#944936]',
        className,
      )}
      {...props}
    />
  );
}

export function Input({ className, ...props }: InputHTMLAttributes<HTMLInputElement>) {
  return (
    <input
      className={clsx(
        'w-full rounded-2xl border border-ink/10 bg-white px-4 py-3 text-sm text-ink outline-none transition focus:border-fairway',
        className,
      )}
      {...props}
    />
  );
}

export function Select({
  className,
  style,
  ...props
}: SelectHTMLAttributes<HTMLSelectElement>) {
  return (
    <select
      className={clsx(
        'w-full appearance-none rounded-2xl border border-ink/10 bg-white px-4 py-3 pr-10 text-sm text-ink outline-none transition focus:border-fairway',
        className,
      )}
      style={{
        backgroundImage:
          "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 12 12'%3E%3Cpath d='M2.25 4.5L6 8.25L9.75 4.5' fill='none' stroke='%23555' stroke-width='1.5' stroke-linecap='round' stroke-linejoin='round'/%3E%3C/svg%3E\")",
        backgroundPosition: 'right 0.75rem center',
        backgroundRepeat: 'no-repeat',
        backgroundSize: '12px 12px',
        ...style,
      }}
      {...props}
    />
  );
}

export function TextArea({ className, ...props }: TextareaHTMLAttributes<HTMLTextAreaElement>) {
  return (
    <textarea
      className={clsx(
        'w-full rounded-3xl border border-ink/10 bg-white px-4 py-3 text-sm text-ink outline-none transition focus:border-fairway',
        className,
      )}
      {...props}
    />
  );
}

export function Badge({
  children,
  tone = 'default',
}: PropsWithChildren<{ tone?: 'default' | 'success' | 'warning' }>) {
  return (
    <span
      className={clsx(
        'inline-flex rounded-full px-3 py-1 text-xs font-semibold',
        tone === 'default' && 'bg-ink/5 text-ink/70',
        tone === 'success' && 'bg-fairway/10 text-fairway',
        tone === 'warning' && 'bg-gold/20 text-[#8b651f]',
      )}
    >
      {children}
    </span>
  );
}
