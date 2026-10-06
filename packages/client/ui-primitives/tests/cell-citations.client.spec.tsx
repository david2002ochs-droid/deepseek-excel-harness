// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest'
import { cleanup, fireEvent, render } from '@testing-library/react'
import { MarkdownText } from './markdown-test-components.tsx'
import { isCellCitation } from '../src/markdown/cell-citations.tsx'

afterEach(cleanup)

describe('workbook-bound cell citations', () => {
  it('renders settled prose citations as buttons and opens the exact address', () => {
    const open = vi.fn()
    const view = render(<MarkdownText text="See [[cite:Sheet1!A1:B3]] and [[cite:'Revenue 2027'!B4:D8]]." cellCitations={{ open }} />)
    expect([...view.container.querySelectorAll('[data-cell-citation]')].map(button => button.textContent))
      .toEqual(['Sheet1!A1:B3', "'Revenue 2027'!B4:D8"])
    fireEvent.click(view.getByRole('button', { name: "'Revenue 2027'!B4:D8" }))
    expect(open).toHaveBeenCalledWith("'Revenue 2027'!B4:D8")
  })

  it('keeps ordinary clients, streaming replies, code, links, and malformed citations inert', () => {
    const owner = { open: vi.fn() }
    const ordinary = render(<MarkdownText text="[[cite:Sheet1!A1]]" />)
    expect(ordinary.container.querySelector('[data-cell-citation]')).toBeNull()
    const streaming = render(<MarkdownText text="[[cite:Sheet1!A1]]" streaming cellCitations={owner} />)
    expect(streaming.container.querySelector('[data-cell-citation]')).toBeNull()
    const code = render(<MarkdownText cellCitations={owner} text={[
      '`[[cite:Sheet1!A1]]`', '', '```text', '[[cite:Sheet1!A1]]', '```', '',
      '[[[cite:Sheet1!A1]]](https://example.com)', '', '[[cite:[Other.xlsx]Sheet1!A1]]',
      '[[cite:Sheet1!B3:A1]] [[cite:https://example.com]] [[cite:Sheet1!XFE1]]',
    ].join('\n')} />)
    expect(code.container.querySelector('[data-cell-citation]')).toBeNull()
    expect(code.container.textContent).toContain('[[cite:Sheet1!A1]]')
  })

  it.each(['Sheet1!A1', "'Revenue 2027'!$B$4:$D$8", "'O''Brien'!A1", "'工作表'!A1", 'Sheet1!XFD1048576'])
  ('accepts bounded in-workbook A1 target %s', (address) => { expect(isCellCitation(address)).toBe(true) })

  it.each(['Sheet1!A0', 'Sheet1!XFE1', 'Sheet1!A1048577', 'Sheet1!B2:A1', 'Sheet1!A:A',
    'Sheet1!1:2', 'Sheet1!A1,B3', '[Other.xlsx]Sheet1!A1', 'https://example.com', 'Sheet1!A1()', "'bad/sheet'!A1"])
  ('leaves unsafe or unbounded target inert %s', (address) => { expect(isCellCitation(address)).toBe(false) })
})
